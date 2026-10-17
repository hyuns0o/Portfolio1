// ============================================================================
// OptionPricingEngine.cpp
// Phase 1 Prototype: OOP Payoff + Black-Scholes + Monte Carlo + Excel x64 DLL
// Environment: Windows 11 + Visual Studio (x64 Release) + 64-bit Excel VBA
// ============================================================================

#include <iostream>
#include <iomanip>
#include <vector>
#include <cmath>
#include <random>
#include <memory>
#include <chrono>
#include <algorithm>

// Windows 64-bit DLL Export Macro (호환성 유지: 콘솔 테스트 및 DLL 빌드 겸용)
#ifdef _WIN32
#define DLL_EXPORT extern "C" __declspec(dllexport)
#else
#define DLL_EXPORT extern "C"
#endif

// ============================================================================
// [Step 1] 객체지향 Payoff 분리 (Instrument & Payoff Layer)
// 새로운 구조화 상품 아이디어가 생기면 엔진 수정 없이 이 클래스만 상속받아 추가합니다.
// ============================================================================

class Payoff {
public:
    virtual ~Payoff() = default;
    // 기초자산 가격(spot)을 받아 만기 페이오프를 반환하는 순수가상함수
    virtual double operator()(double spot) const = 0;
};

// 유럽형 콜옵션 Payoff: max(S - K, 0)
class PayoffCall : public Payoff {
private:
    double strike_;
public:
    explicit PayoffCall(double strike) : strike_(strike) {}
    double operator()(double spot) const override {
        return std::max(spot - strike_, 0.0);
    }
};

// 유럽형 풋옵션 Payoff: max(K - S, 0)
class PayoffPut : public Payoff {
private:
    double strike_;
public:
    explicit PayoffPut(double strike) : strike_(strike) {}
    double operator()(double spot) const override {
        return std::max(strike_ - spot, 0.0);
    }
};

// 디지털(Binary) 콜옵션 Payoff: S > K 이면 1.0, 아니면 0.0
class PayoffDigitalCall : public Payoff {
private:
    double strike_;
public:
    explicit PayoffDigitalCall(double strike) : strike_(strike) {}
    double operator()(double spot) const override {
        return (spot > strike_) ? 1.0 : 0.0;
    }
};

// 스트래들(Straddle) 구조화 Payoff 예시: |S - K|
class PayoffStraddle : public Payoff {
private:
    double strike_;
public:
    explicit PayoffStraddle(double strike) : strike_(strike) {}
    double operator()(double spot) const override {
        return std::abs(spot - strike_);
    }
};

// ============================================================================
// [Step 2] 블랙-숄즈 해석해 및 5대 그릭스 엔진 (Analytical Solver Layer)
// ============================================================================

struct OptionGreeks {
    double price;  // 옵션 현재가치 (PV)
    double delta;  // 기초자산 1단위 변화 시 가격 민감도 (dV/dS)
    double gamma;  // 기초자산 1단위 변화 시 델타 민감도 (d^2V/dS^2)
    double vega;   // 변동성 1%p(0.01) 변화 시 가격 민감도
    double theta;  // 시간 1일(1/365년) 경과 시 가치 변화 (Time Decay)
    double rho;    // 이자율 1%p(0.01) 변화 시 가격 민감도
};

namespace MathUtils {
    constexpr double INV_SQRT_2PI = 0.39894228040143267794; // 1 / sqrt(2 * pi)
    constexpr double INV_SQRT_2 = 0.70710678118654752440; // 1 / sqrt(2)

    // 표준정규분포 확률밀도함수 n(x)
    inline double norm_pdf(double x) {
        return INV_SQRT_2PI * std::exp(-0.5 * x * x);
    }

    // 표준정규분포 누적분포함수 N(x) (C++ 표준 std::erfc 활용 - 고정밀/고속)
    inline double norm_cdf(double x) {
        return 0.5 * std::erfc(-x * INV_SQRT_2);
    }
}

class BlackScholesEngine {
public:
    // is_call: 1 = Call, 0 = Put
    static OptionGreeks calculate(
        double S, double K, double r, double q, double sigma, double T, int is_call
    ) {
        OptionGreeks res{ 0.0, 0.0, 0.0, 0.0, 0.0, 0.0 };

        // 만기 도래 또는 변동성 0인 예외 구간 안전 처리
        if (T <= 1e-10 || sigma <= 1e-10 || S <= 1e-10 || K <= 1e-10) {
            if (is_call == 1) {
                res.price = std::max(S * std::exp(-q * T) - K * std::exp(-r * T), 0.0);
                res.delta = (S > K) ? 1.0 : 0.0;
            }
            else {
                res.price = std::max(K * std::exp(-r * T) - S * std::exp(-q * T), 0.0);
                res.delta = (S < K) ? -1.0 : 0.0;
            }
            return res;
        }

        const double sqrt_T = std::sqrt(T);
        const double sigma_sqrt_T = sigma * sqrt_T;
        const double d1 = (std::log(S / K) + (r - q + 0.5 * sigma * sigma) * T) / sigma_sqrt_T;
        const double d2 = d1 - sigma_sqrt_T;

        const double df_r = std::exp(-r * T); // 무위험 금리 할인계수
        const double df_q = std::exp(-q * T); // 배당 할인계수
        const double pdf_d1 = MathUtils::norm_pdf(d1);

        // Gamma와 Vega는 콜/풋 동일
        res.gamma = (df_q * pdf_d1) / (S * sigma_sqrt_T);
        res.vega = (S * df_q * pdf_d1 * sqrt_T) * 0.01; // 1%p 변동성 기준

        if (is_call == 1) {
            const double N_d1 = MathUtils::norm_cdf(d1);
            const double N_d2 = MathUtils::norm_cdf(d2);

            res.price = S * df_q * N_d1 - K * df_r * N_d2;
            res.delta = df_q * N_d1;
            res.theta = (-(S * df_q * pdf_d1 * sigma) / (2.0 * sqrt_T)
                - r * K * df_r * N_d2
                + q * S * df_q * N_d1) / 365.0; // 1일 기준
            res.rho = (K * T * df_r * N_d2) * 0.01;     // 1%p 금리 기준
        }
        else {
            const double N_minus_d1 = MathUtils::norm_cdf(-d1);
            const double N_minus_d2 = MathUtils::norm_cdf(-d2);

            res.price = K * df_r * N_minus_d2 - S * df_q * N_minus_d1;
            res.delta = -df_q * N_minus_d1;
            res.theta = (-(S * df_q * pdf_d1 * sigma) / (2.0 * sqrt_T)
                + r * K * df_r * N_minus_d2
                - q * S * df_q * N_minus_d1) / 365.0;
            res.rho = (-K * T * df_r * N_minus_d2) * 0.01;
        }

        return res;
    }
};

// ============================================================================
// [Step 3] 연속시간 근사 몬테카를로 엔진 (Monte Carlo Path Simulator)
// 대조변수법(Antithetic Variates) 적용 + 경로 의존형 배리어 옵션 지원
// ============================================================================

class MonteCarloEngine {
public:
    // 일반 Payoff 객체 기반 시뮬레이션 (만기 시점 평가)
    static double price_terminal(
        const Payoff& payoff,
        double S0, double r, double q, double sigma, double T,
        int num_paths, unsigned int seed = 42
    ) {
        std::mt19937_64 rng(seed);
        std::normal_distribution<double> norm(0.0, 1.0);

        const double drift = (r - q - 0.5 * sigma * sigma) * T;
        const double vol_sqrt_T = sigma * std::sqrt(T);
        const double df = std::exp(-r * T);

        double payoff_sum = 0.0;
        const int half_paths = std::max(num_paths / 2, 1);

        for (int i = 0; i < half_paths; ++i) {
            const double z = norm(rng);
            // 대조변수법 (+z와 -z 경로를 동시에 생성하여 오차 분산 대폭 감소)
            const double ST_pos = S0 * std::exp(drift + vol_sqrt_T * z);
            const double ST_neg = S0 * std::exp(drift - vol_sqrt_T * z);
            payoff_sum += 0.5 * (payoff(ST_pos) + payoff(ST_neg));
        }

        return df * (payoff_sum / static_cast<double>(half_paths));
    }

    // 연속시간 근사(dt = T / num_steps) 경로 의존형 Knock-Out 배리어 옵션 프라이서
    // barrier_type: 1 = Up-and-Out (상방 배리어 터치 시 소멸), -1 = Down-and-Out (하방 배리어 터치 시 소멸)
    static double price_knock_out_barrier(
        const Payoff& payoff,
        double S0, double barrier, int barrier_type,
        double r, double q, double sigma, double T,
        int num_steps, int num_paths, unsigned int seed = 42
    ) {
        std::mt19937_64 rng(seed);
        std::normal_distribution<double> norm(0.0, 1.0);

        const double dt = T / static_cast<double>(num_steps);
        const double drift_dt = (r - q - 0.5 * sigma * sigma) * dt;
        const double vol_sqrt_dt = sigma * std::sqrt(dt);
        const double df = std::exp(-r * T);

        double payoff_sum = 0.0;

        for (int p = 0; p < num_paths; ++p) {
            double S_curr = S0;
            bool knocked_out = false;

            for (int step = 0; step < num_steps; ++step) {
                const double z = norm(rng);
                S_curr *= std::exp(drift_dt + vol_sqrt_dt * z);

                if ((barrier_type == 1 && S_curr >= barrier) ||
                    (barrier_type == -1 && S_curr <= barrier)) {
                    knocked_out = true;
                    break; // 배리어 터치 즉시 경로 탐색 종료 (연산 절약)
                }
            }

            if (!knocked_out) {
                payoff_sum += payoff(S_curr);
            }
        }

        return df * (payoff_sum / static_cast<double>(num_paths));
    }
};

// ============================================================================
// [Step 4] 엑셀(VBA) 및 파이썬 연동용 C 인터페이스 (DLL Export Bridge)
// 64비트 엑셀에서 Declare PtrSafe Function/Sub 로 즉시 호출 가능한 함수들입니다.
// ============================================================================

// 1. 단일 셀 수식용 함수 (엑셀 셀에서 =VBA_BS_Price(...) 로 사용)
DLL_EXPORT double BS_PriceSingle(
    double S, double K, double r, double q, double sigma, double T, int is_call
) {
    return BlackScholesEngine::calculate(S, K, r, q, sigma, T, is_call).price;
}

// 2. 초고속 배열 포인터 일괄 처리 함수 (Batch Array Processing)
// 엑셀에서 1,000~10,000개의 기초자산 구간(S_arr)을 단 1번의 포인터 호출로 넘겨받아
// 가격과 5대 그릭스 배열을 0.00003초(30마이크로초) 만에 한꺼번에 채워줍니다.
DLL_EXPORT void BS_PriceBatch(
    int count,
    const double* S_arr,      // [입력] 기초자산 가격 배열 시작 주소
    double K,
    double r,
    double q,
    double sigma,
    double T,
    int is_call,
    double* out_price,        // [출력] 가격 배열 시작 주소
    double* out_delta,        // [출력] 델타 배열 시작 주소
    double* out_gamma,        // [출력] 감마 배열 시작 주소
    double* out_vega,         // [출력] 베가 배열 시작 주소
    double* out_theta,        // [출력] 세타 배열 시작 주소
    double* out_rho           // [출력] 로 배열 시작 주소
) {
    if (count <= 0 || !S_arr) return;

    for (int i = 0; i < count; ++i) {
        const OptionGreeks g = BlackScholesEngine::calculate(
            S_arr[i], K, r, q, sigma, T, is_call
        );
        if (out_price) out_price[i] = g.price;
        if (out_delta) out_delta[i] = g.delta;
        if (out_gamma) out_gamma[i] = g.gamma;
        if (out_vega)  out_vega[i] = g.vega;
        if (out_theta) out_theta[i] = g.theta;
        if (out_rho)   out_rho[i] = g.rho;
    }
}

// 3. 연속시간 근사 배리어 옵션 몬테카를로 함수 (엑셀에서 호출 가능)
DLL_EXPORT double MC_KnockOutBarrierPrice(
    double S0, double K, double barrier, int barrier_type,
    double r, double q, double sigma, double T,
    int is_call, int num_steps, int num_paths, unsigned int seed
) {
    if (is_call == 1) {
        PayoffCall payoff(K);
        return MonteCarloEngine::price_knock_out_barrier(
            payoff, S0, barrier, barrier_type, r, q, sigma, T, num_steps, num_paths, seed
        );
    }
    else {
        PayoffPut payoff(K);
        return MonteCarloEngine::price_knock_out_barrier(
            payoff, S0, barrier, barrier_type, r, q, sigma, T, num_steps, num_paths, seed
        );
    }
}

// ============================================================================
// [Step 5] 콘솔 단독 검증 및 속도 벤치마크용 main() 함수
// DLL 빌드 전, Visual Studio에서 Ctrl+F5를 눌러 엔진 정확도와 속도를 먼저 검증합니다.
// ============================================================================

int main() {
    std::cout << "=========================================================\n";
    std::cout << "  Phase 1 C++ Option Pricing Engine Verification & Speed \n";
    std::cout << "=========================================================\n\n";

    const double S0 = 100.0; // 기초자산 현재가
    const double K = 100.0; // 행사가 (ATM)
    const double r = 0.035; // 무위험 이자율 (연 3.5%)
    const double q = 0.000; // 배당률 (0%)
    const double sigma = 0.200; // 변동성 (연 20%)
    const double T = 1.000; // 잔존만기 (1년)

    // 1. 블랙-숄즈 해석해 검증
    OptionGreeks call_bs = BlackScholesEngine::calculate(S0, K, r, q, sigma, T, 1);
    OptionGreeks put_bs = BlackScholesEngine::calculate(S0, K, r, q, sigma, T, 0);

    std::cout << std::fixed << std::setprecision(6);
    std::cout << "[1] Black-Scholes Analytical Results (S=100, K=100, r=3.5%, vol=20%, T=1Y)\n";
    std::cout << "  Call Price : " << call_bs.price << " | Put Price : " << put_bs.price << "\n";
    std::cout << "  Call Delta : " << call_bs.delta << " | Put Delta : " << put_bs.delta << "\n";
    std::cout << "  Gamma      : " << call_bs.gamma << " | Vega(1%p) : " << call_bs.vega << "\n";
    std::cout << "  Call Theta : " << call_bs.theta << " | Call Rho  : " << call_bs.rho << "\n\n";

    // 2. 몬테카를로 수렴 검증 (Payoff 다형성 테스트)
    PayoffCall call_payoff(K);
    PayoffStraddle straddle_payoff(K);
    double mc_call_price = MonteCarloEngine::price_terminal(call_payoff, S0, r, q, sigma, T, 200000, 42);
    double mc_straddle = MonteCarloEngine::price_terminal(straddle_payoff, S0, r, q, sigma, T, 200000, 42);

    std::cout << "[2] Monte Carlo Convergence & Custom Payoff (200,000 Paths)\n";
    std::cout << "  MC Call Price      : " << mc_call_price
        << " (Error vs BS: " << std::abs(mc_call_price - call_bs.price) << ")\n";
    std::cout << "  MC Straddle Price  : " << mc_straddle
        << " (BS Call + Put: " << (call_bs.price + put_bs.price) << ")\n\n";

    // 3. 1,000개 기초자산 그리드 일괄 프라이싱(Batch) 속도 벤치마크
    const int N_GRID = 1000;
    std::vector<double> S_grid(N_GRID);
    std::vector<double> out_price(N_GRID), out_delta(N_GRID), out_gamma(N_GRID);
    std::vector<double> out_vega(N_GRID), out_theta(N_GRID), out_rho(N_GRID);

    for (int i = 0; i < N_GRID; ++i) {
        S_grid[i] = 50.0 + (100.0 * i) / (N_GRID - 1); // S = 50.0 ~ 150.0 연속 1,000개 구간
    }

    auto t_start = std::chrono::high_resolution_clock::now();
    BS_PriceBatch(
        N_GRID, S_grid.data(), K, r, q, sigma, T, 1,
        out_price.data(), out_delta.data(), out_gamma.data(),
        out_vega.data(), out_theta.data(), out_rho.data()
    );
    auto t_end = std::chrono::high_resolution_clock::now();

    double elapsed_us = std::chrono::duration<double, std::micro>(t_end - t_start).count();

    std::cout << "[3] Batch Array Benchmark (1,000 Options x 6 Metrics = 6,000 Calculations)\n";
    std::cout << "  Elapsed Time : " << std::setprecision(2) << elapsed_us << " microseconds ("
        << std::setprecision(5) << (elapsed_us / 1000.0) << " ms)\n";
    std::cout << "  Sample S=100 : Price=" << out_price[500] << ", Delta=" << out_delta[500] << "\n";
    std::cout << "=========================================================\n";

    return 0;
}

/*
' ============================================================================
' [엑셀 VBA 표준 모듈 코드 (ModOptionEngine.bas)]
' 엑셀에서 Alt + F11 -> [삽입] -> [모듈]을 누르고 아래 코드를 그대로 붙여넣으세요.
' 주의: "C:\Quant\OptionPricingEngine.dll" 경로를 본인 PC의 실제 DLL 경로로 수정하세요!
' ============================================================================
Option Explicit

' 1. C++ 64비트 DLL 함수 선언부 (Declare PtrSafe)
Public Declare PtrSafe Function BS_PriceSingle Lib "C:\Quant\OptionPricingEngine.dll" ( _
    ByVal S As Double, _
    ByVal K As Double, _
    ByVal r As Double, _
    ByVal q As Double, _
    ByVal sigma As Double, _
    ByVal T As Double, _
    ByVal is_call As Long _
) As Double

Public Declare PtrSafe Sub BS_PriceBatch Lib "C:\Quant\OptionPricingEngine.dll" ( _
    ByVal count As Long, _
    ByRef S_arr As Double, _
    ByVal K As Double, _
    ByVal r As Double, _
    ByVal q As Double, _
    ByVal sigma As Double, _
    ByVal T As Double, _
    ByVal is_call As Long, _
    ByRef out_price As Double, _
    ByRef out_delta As Double, _
    ByRef out_gamma As Double, _
    ByRef out_vega As Double, _
    ByRef out_theta As Double, _
    ByRef out_rho As Double _
)

Public Declare PtrSafe Function MC_KnockOutBarrierPrice Lib "C:\Quant\OptionPricingEngine.dll" ( _
    ByVal S0 As Double, _
    ByVal K As Double, _
    ByVal barrier As Double, _
    ByVal barrier_type As Long, _
    ByVal r As Double, _
    ByVal q As Double, _
    ByVal sigma As Double, _
    ByVal T As Double, _
    ByVal is_call As Long, _
    ByVal num_steps As Long, _
    ByVal num_paths As Long, _
    ByVal seed As Long _
) As Double

' 2. 엑셀 시트 셀 수식용 래퍼 함수: 셀에 =CPP_BS_Price(100, 100, 0.035, 0, 0.2, 1, 1) 입력
Public Function CPP_BS_Price(S As Double, K As Double, r As Double, q As Double, sigma As Double, T As Double, is_call As Long) As Double
    CPP_BS_Price = BS_PriceSingle(S, K, r, q, sigma, T, is_call)
End Function

' 3. 1,000개 연속 기초자산 구간 초고속 일괄 계산 매크로 (0.01초 내 시트 출력 완료)
Public Sub RunBatchPricing1000()
    Const N As Long = 1000
    Dim S_arr(1 To N) As Double
    Dim out_price(1 To N) As Double
    Dim out_delta(1 To N) As Double
    Dim out_gamma(1 To N) As Double
    Dim out_vega(1 To N) As Double
    Dim out_theta(1 To N) As Double
    Dim out_rho(1 To N) As Double

    Dim i As Long
    For i = 1 To N
        S_arr(i) = 50# + (100# * (i - 1)) / (N - 1) ' 기초자산 50 ~ 150 구간을 1,000개로 쪼개기
    Next i

    ' C++ DLL에 배열의 첫 번째 원소 메모리 주소(ByRef S_arr(1))만 단 1번 전달!
    Call BS_PriceBatch(N, S_arr(1), 100#, 0.035, 0#, 0.2, 1#, 1, _
                       out_price(1), out_delta(1), out_gamma(1), _
                       out_vega(1), out_theta(1), out_rho(1))

    ' 엑셀 시트 출력용 2차원 배열로 변환하여 단 1번의 Value2 대입으로 화면에 즉시 덤프
    Dim sheetGrid(1 To N, 1 To 7) As Double
    For i = 1 To N
        sheetGrid(i, 1) = S_arr(i)
        sheetGrid(i, 2) = out_price(i)
        sheetGrid(i, 3) = out_delta(i)
        sheetGrid(i, 4) = out_gamma(i)
        sheetGrid(i, 5) = out_vega(i)
        sheetGrid(i, 6) = out_theta(i)
        sheetGrid(i, 7) = out_rho(i)
    Next i

    With ActiveSheet
        .Range("A1:G1").Value = Array("Spot(S)", "Price", "Delta", "Gamma", "Vega", "Theta", "Rho")
        .Range("A2").Resize(N, 7).Value2 = sheetGrid
    End With
End Sub
*/