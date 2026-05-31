#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Расчёт экономической эффективности стратегии продвижения Face2 (ВКР, гл. 3).
Горизонт: 2026 (год 0) — 2030 (год 4). Ставка дисконтирования r = 16% годовых.
Все суммы — тыс. руб., если не указано иное.

Запуск:    python3 _calc_chapter3.py
Результат: печать сводной таблицы и детализации в stdout.
"""

# ---------------- 1. Базовые параметры ----------------
RATE = 0.16  # ставка дисконтирования
HORIZON_YEARS = [2027, 2028, 2029, 2030]
PREV_YEAR = {2027: 2026, 2028: 2027, 2029: 2028, 2030: 2029}

# ---------------- 2. Инвестиции и OPEX ----------------
I0_BREAKDOWN = {
    "site":      3_500,
    "certif":    5_000,
    "brand":     1_500,
    "content":   1_500,
    "pilot_kit": 1_000,
}
I0 = sum(I0_BREAKDOWN.values())  # 12 500

OPEX_2027 = {"smm": 2_400, "pr": 3_600, "perf": 3_600, "events": 4_500,
             "partners": 3_000, "content": 2_400, "team": 5_400}

def _opex_2028():
    return round(OPEX_2027["smm"]*1.08 + OPEX_2027["pr"]*1.08
                 + OPEX_2027["perf"]*1.25 + OPEX_2027["events"]*1.08
                 + OPEX_2027["partners"]*1.25 + OPEX_2027["content"]*1.08
                 + OPEX_2027["team"]*1.10)

OPEX_BY_YEAR = {
    2027: sum(OPEX_2027.values()),
    2028: _opex_2028(),
}
OPEX_BY_YEAR[2029] = round(OPEX_BY_YEAR[2028] * 1.08)
OPEX_BY_YEAR[2030] = round(OPEX_BY_YEAR[2029] * 1.06)

# ---------------- 3. Параметры выручки ----------------
KBS_DEAL, KBS_SUPPORT = 25_000, 4_000
F2P_CONNECT, F2P_FEE_M = 80, 25
F2PASS_LIC, F2PASS_TERM_PRICE, F2PASS_FEE_M = 1_200, 120, 30
F2PASS_TERMS_PER_OBJ = 6
F2CI_CONNECT, F2CI_FEE_M = 150, 20

MARGIN_LIC, MARGIN_INT = 0.65, 0.45

SCENARIOS = {
    "Pessimistic": {
        "KBS":    {2027: 0, 2028: 1, 2029: 2,  2030: 3},
        "F2P":    {2027: 3, 2028: 6, 2029: 10, 2030: 15},
        "F2Pass": {2027: 1, 2028: 2, 2029: 4,  2030: 6},
        "F2Ci":   {2027: 1, 2028: 2, 2029: 4,  2030: 6},
    },
    "Realistic": {
        "KBS":    {2027: 1, 2028: 2,  2029: 4,  2030: 6},
        "F2P":    {2027: 6, 2028: 14, 2029: 25, 2030: 40},
        "F2Pass": {2027: 2, 2028: 5,  2029: 9,  2030: 14},
        "F2Ci":   {2027: 2, 2028: 5,  2029: 9,  2030: 14},
    },
    "Optimistic": {
        "KBS":    {2027: 2,  2028: 4,  2029: 7,  2030: 10},
        "F2P":    {2027: 10, 2028: 24, 2029: 42, 2030: 65},
        "F2Pass": {2027: 4,  2028: 9,  2029: 16, 2030: 25},
        "F2Ci":   {2027: 4,  2028: 9,  2029: 16, 2030: 25},
    },
}


# ---------------- 4. Расчёт выручки за год ----------------
def calc_year(scn, year, prev_year):
    """Вычисляет выручку, валовую прибыль и компоненты за год year."""
    def diff(prod):
        total = scn[prod][year]
        prev = scn[prod][prev_year] if prev_year >= 2027 else 0
        return total, prev, total - prev

    kbs_tot, kbs_prev, kbs_new = diff("KBS")
    kbs_lic = kbs_new * KBS_DEAL * 0.6 + kbs_tot * KBS_SUPPORT
    kbs_int = kbs_new * KBS_DEAL * 0.4

    f2p_tot, f2p_prev, f2p_new = diff("F2P")
    f2p_lic = f2p_new * (12 * F2P_FEE_M) + f2p_prev * (12 * F2P_FEE_M)
    f2p_int = f2p_new * F2P_CONNECT

    f2pass_tot, f2pass_prev, f2pass_new = diff("F2Pass")
    f2pass_lic = f2pass_new * (F2PASS_LIC + 12 * F2PASS_FEE_M) + f2pass_prev * (12 * F2PASS_FEE_M)
    f2pass_int = f2pass_new * (F2PASS_TERMS_PER_OBJ * F2PASS_TERM_PRICE)

    f2ci_tot, f2ci_prev, f2ci_new = diff("F2Ci")
    f2ci_lic = f2ci_new * (12 * F2CI_FEE_M) + f2ci_prev * (12 * F2CI_FEE_M)
    f2ci_int = f2ci_new * F2CI_CONNECT

    rev_lic = kbs_lic + f2p_lic + f2pass_lic + f2ci_lic
    rev_int = kbs_int + f2p_int + f2pass_int + f2ci_int
    return {
        "rev_lic": rev_lic,
        "rev_int": rev_int,
        "rev_total": rev_lic + rev_int,
        "gp": rev_lic * MARGIN_LIC + rev_int * MARGIN_INT,
        "components": {
            "KBS_lic": kbs_lic, "KBS_int": kbs_int,
            "F2P_lic": f2p_lic, "F2P_int": f2p_int,
            "F2Pass_lic": f2pass_lic, "F2Pass_int": f2pass_int,
            "F2Ci_lic": f2ci_lic, "F2Ci_int": f2ci_int,
        },
        "new": {"KBS": kbs_new, "F2P": f2p_new,
                "F2Pass": f2pass_new, "F2Ci": f2ci_new},
    }


# ---------------- 5. Метрики эффективности ----------------
def npv(cfs, rate):
    return sum(cf / (1 + rate) ** t for t, cf in enumerate(cfs))


def irr(cfs, low=-0.99, high=10.0, tol=1e-7, max_iter=300):
    f_low, f_high = npv(cfs, low), npv(cfs, high)
    if f_low * f_high > 0:
        return None
    for _ in range(max_iter):
        mid = (low + high) / 2
        f_mid = npv(cfs, mid)
        if abs(f_mid) < tol:
            return mid
        if f_low * f_mid < 0:
            high, f_high = mid, f_mid
        else:
            low, f_low = mid, f_mid
    return (low + high) / 2


def discounted_payback(cfs, rate):
    cum_prev, cum = 0.0, cfs[0]
    for t in range(1, len(cfs)):
        disc = cfs[t] / (1 + rate) ** t
        cum_prev = cum
        cum += disc
        if cum_prev < 0 <= cum:
            return (t - 1) + (-cum_prev) / disc
    return None


def simple_payback(cfs):
    cum_prev, cum = 0.0, cfs[0]
    for t in range(1, len(cfs)):
        cum_prev = cum
        cum += cfs[t]
        if cum_prev < 0 <= cum:
            return (t - 1) + (-cum_prev) / cfs[t]
    return None


def profitability_index(cfs, rate):
    pv_pos = sum(cf / (1 + rate) ** t for t, cf in enumerate(cfs)
                 if t > 0 and cf > 0)
    pv_neg = sum(-cf / (1 + rate) ** t for t, cf in enumerate(cfs) if cf < 0)
    return pv_pos / pv_neg


# ---------------- 6. Прогон сценариев ----------------
def run_scenarios():
    out = {}
    for name, scn in SCENARIOS.items():
        yres = {y: calc_year(scn, y, PREV_YEAR[y]) for y in HORIZON_YEARS}
        cfs = [-I0]
        for y in HORIZON_YEARS:
            cfs.append(yres[y]["gp"] - OPEX_BY_YEAR[y])
        out[name] = {
            "yres": yres, "cfs": cfs,
            "NPV": npv(cfs, RATE),
            "PI": profitability_index(cfs, RATE),
            "IRR": irr(cfs),
            "DPP": discounted_payback(cfs, RATE),
            "PP": simple_payback(cfs),
            "ROI_simple": (sum(cfs[1:]) / -cfs[0]) * 100,
        }
    return out


# ---------------- 7. Печать ----------------
def _safe(v, fmt):
    return "не окуп." if v is None else fmt.format(v)


def print_summary(results):
    print(f"Стартовые инвестиции I0 (2026): {I0:,} тыс. руб.")
    print("OPEX по годам (тыс. руб.):", end=" ")
    print(" | ".join(f"{y}={OPEX_BY_YEAR[y]:,}" for y in HORIZON_YEARS))
    print()
    print("=" * 82)
    print(f"{'Показатель (тыс. руб.)':<40}{'Π':>14}{'Р':>14}{'О':>14}")
    print("=" * 82)
    keys = ("Pessimistic", "Realistic", "Optimistic")
    for y in HORIZON_YEARS:
        print(f"{'Выручка ' + str(y):<40}"
              + "".join(f"{results[s]['yres'][y]['rev_total']:>14,.0f}" for s in keys))
    for y in HORIZON_YEARS:
        print(f"{'Валовая прибыль ' + str(y):<40}"
              + "".join(f"{results[s]['yres'][y]['gp']:>14,.0f}" for s in keys))
    print("-" * 82)
    print(f"{'CF0 (2026, инвест.)':<40}"
          + "".join(f"{results[s]['cfs'][0]:>14,.0f}" for s in keys))
    for i, y in enumerate(HORIZON_YEARS, 1):
        print(f"{'CF' + str(i) + ' (' + str(y) + ')':<40}"
              + "".join(f"{results[s]['cfs'][i]:>14,.0f}" for s in keys))
    print("-" * 82)
    print(f"{'NPV (ЧДД), r=16%':<40}"
          + "".join(f"{results[s]['NPV']:>14,.0f}" for s in keys))
    print(f"{'PI (Индекс прибыльности)':<40}"
          + "".join(f"{results[s]['PI']:>14.2f}" for s in keys))
    print(f"{'IRR (ВНД), %':<40}"
          + "".join((f"{results[s]['IRR'] * 100:>14.1f}"
                     if results[s]['IRR'] is not None else f"{'—':>14}")
                    for s in keys))
    print(f"{'DPP, лет (дисконт.)':<40}"
          + "".join(f"{_safe(results[s]['DPP'], '{:.2f}'):>14}" for s in keys))
    print(f"{'PP, лет (простой)':<40}"
          + "".join(f"{_safe(results[s]['PP'], '{:.2f}'):>14}" for s in keys))
    print(f"{'ROI простой за 4 года, %':<40}"
          + "".join(f"{results[s]['ROI_simple']:>14.1f}" for s in keys))
    print("=" * 82)

    # Чувствительность NPV к ставке
    print("\nЧувствительность NPV (реалист.) к ставке r:")
    for rr in [0.10, 0.13, 0.145, 0.16, 0.18, 0.20, 0.25, 0.30]:
        print(f"  r = {rr * 100:>5.1f}%  →  NPV = "
              f"{npv(results['Realistic']['cfs'], rr):>12,.0f} тыс. руб.")

    # Терминальная стоимость
    g = 0.04
    cf4 = results["Realistic"]["cfs"][4]
    tv = cf4 * (1 + g) / (RATE - g)
    pv_tv = tv / (1 + RATE) ** 4
    print(f"\nТерминальная стоимость (g=4%):")
    print(f"  CF4 (2030) = {cf4:,.0f}")
    print(f"  TV(2030) = CF4·(1+g)/(r−g) = {tv:,.0f}")
    print(f"  PV(TV) = {pv_tv:,.0f}")
    print(f"  NPV + PV(TV) = {results['Realistic']['NPV'] + pv_tv:,.0f}")


if __name__ == "__main__":
    print_summary(run_scenarios())
