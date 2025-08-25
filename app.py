# app.py — EV Charger Business Analyzer (Web / Streamlit)
# ---------------------------------------------------------
# - Mirrors the desktop Tkinter tool's logic in a browser UI
# - Tiered subsidy or manual per-unit; manual totals override
# - Manual constants for base fee (per kW), comms, mgmt, and capacity (kW/charger)
# - "Grant-surplus" case → simple mode (annual/monthly profit focus)
# - Export Excel with two sheets: 분석 요약 (wide, column-per-run) + History (append rows)
# - Optionally upload an existing master Excel to append to (no server persistence)
#
# Run locally:
#   pip install streamlit pandas openpyxl
#   streamlit run app.py
#
import io
from datetime import datetime

import pandas as pd
import streamlit as st

# ----------------------
# Utility functions
# ----------------------

def calc_tiered_subsidy(n: int) -> int:
    """Tiered total subsidy: 1→2,200,000; 2–5→2,000,000 ea; 6+→1,800,000 ea."""
    if n <= 0:
        return 0
    total = 0
    total += 2_200_000 if n >= 1 else 0
    if n >= 2:
        cnt_2_5 = min(n - 1, 4)
        total += cnt_2_5 * 2_000_000
    if n >= 6:
        total += (n - 5) * 1_800_000
    return total


def append_summary_as_column(existing_wide: pd.DataFrame | None,
                             df_two_col: pd.DataFrame,
                             col_name: str) -> pd.DataFrame:
    """Create/append a wide summary table where each run becomes a new column.
    df_two_col must have columns ['항목','값'].
    """
    if existing_wide is None or '항목' not in existing_wide.columns:
        return df_two_col.rename(columns={'값': col_name}).copy()
    wide = existing_wide.copy()
    merged = pd.merge(wide, df_two_col, on='항목', how='outer')
    merged = merged.rename(columns={'값': col_name})
    # keep the latest order of items
    order_map = {v: i for i, v in enumerate(df_two_col['항목'].tolist())}
    merged['__ord__'] = merged['항목'].map(order_map).fillna(9999).astype(int)
    merged = merged.sort_values('__ord__').drop(columns='__ord__')
    return merged


def build_summary_two_col(site_name: str, num_chargers: int,
                          cost_install_ops: float, cost_hw_materials: float, per_unit_gross_cost: float,
                          investment_per_charger: float, initial_investment: float,
                          monthly_revenue_total: float, monthly_cost_total: float, monthly_profit_total: float,
                          profit_1y_per_charger: float, payback_period_months: float | None, roi_per_charger: float | None,
                          total_subsidy: float, avg_subsidy_per_unit: float, total_gross_capex: float,
                          is_tiered: bool,
                          base_fee_per_kw: float, comms_per_charger: float, mgmt_per_charger: float, capacity_kw: float,
                          m_base: bool, m_comms: bool, m_mgmt: bool, m_cap: bool,
                          total_capacity_kw: float, invest_per_kw: float, monthly_profit_per_kw: float,
                          m_total_inv_on: bool, m_total_inv_val: float,
                          m_total_sub_on: bool, m_total_sub_val: float,
                          grant_surplus: float, subsidy_coverage_ratio: float | None, gross_payback_months: float | None) -> pd.DataFrame:
    const_mode = f"{'수동' if m_base else '자동'}/{'수동' if m_comms else '자동'}/{'수동' if m_mgmt else '자동'}/{'수동' if m_cap else '자동'}"
    subsidy_mode_label = "계단식" if is_tiered else "1기당 수동"

    items = [
        '현장명', '총 설치 대수',
        '1-1 영업+시공(기당)', '1-2 충전기+부자재(기당)', '기당 총 원가',
        '총 원가(설치비 합계)', '총 보조금(최종)', '총 투자비(최종, 보조금 차감)',
        '총 보조금(수동 입력값)', '총 투자비(수동 입력값)',
        '기당 평균 보조금', '기기당 실투자비', '보조금 방식(1기 기준)', '---',
        '충전기 용량(기당 kW)', '총 충전기 용량(kW)',
        '전력 기본요금(kW·월)', '통신비(기·월)', '관리비(기·월)',
        '상수/스펙 입력 모드(기본/통신/관리/용량)', '---',
        '투자금/총 용량(원/kW)', '월 순이익/총 용량(원/kW·월)', '보조금 잉여(초기 유입 원)', '보조금 커버리지(보조금/총원가)', '총원가 기준 회수기간(개월)', '---',
        '월 총 매출(모델, 30일)', '월 총 비용(모델, 30일)', '월 총 순이익(모델, 30일)', '---',
        '기기당 연간 순이익', '투자비 회수 기간(개월)', '연간 투자수익률(ROI)'
    ]

    vals = [
        site_name, f"{num_chargers} 대",
        f"{cost_install_ops:,.0f} 원", f"{cost_hw_materials:,.0f} 원", f"{per_unit_gross_cost:,.0f} 원",
        f"{total_gross_capex:,.0f} 원", f"{total_subsidy:,.0f} 원", f"{initial_investment:,.0f} 원",
        ("" if not m_total_sub_on else f"{m_total_sub_val:,.0f} 원"),
        ("" if not m_total_inv_on else f"{m_total_inv_val:,.0f} 원"),
        f"{avg_subsidy_per_unit:,.0f} 원", f"{investment_per_charger:,.0f} 원", subsidy_mode_label, "",
        f"{capacity_kw:,.2f} kW", f"{total_capacity_kw:,.2f} kW",
        f"{base_fee_per_kw:,.0f} 원", f"{comms_per_charger:,.0f} 원", f"{mgmt_per_charger:,.0f} 원",
        const_mode, "",
        f"{invest_per_kw:,.0f} 원/kW", f"{monthly_profit_per_kw:,.0f} 원/kW·월",
        f"{grant_surplus:,.0f} 원",
        ("N/A" if subsidy_coverage_ratio is None else f"{subsidy_coverage_ratio:.2f} x"),
        ("∞" if (gross_payback_months is None or gross_payback_months == float('inf')) else f"{gross_payback_months:.1f}"),
        "",
        f"{monthly_revenue_total:,.0f} 원", f"{monthly_cost_total:,.0f} 원", f"{monthly_profit_total:,.0f} 원", "",
        f"{profit_1y_per_charger:,.0f} 원",
        ("회수 불가" if payback_period_months == float('inf') else f"{payback_period_months:.1f} 개월"),
        ("N/A" if roi_per_charger is None else f"{roi_per_charger:.2f} %")
    ]
    return pd.DataFrame({'항목': items, '값': vals})


# ----------------------
# Page Config & Header
# ----------------------
st.set_page_config(page_title="EV Charger Analyzer (Web)", layout="wide")
st.title("⚡ 전기차 충전기 사업성 분석기 — Web")
st.caption("데스크탑 버전의 계산 로직을 그대로 웹에서 사용할 수 있게 구성했습니다.")

# Optional master Excel to append to
with st.expander("옵션: 기존 마스터 엑셀 업로드(있다면 여기에 결과를 이어 붙여 드립니다)"):
    master_file = st.file_uploader("분석 요약/History 시트가 있는 기존 xlsx (선택)", type=["xlsx"])    
    existing_summary = None
    existing_history = None
    if master_file is not None:
        try:
            xls = pd.ExcelFile(master_file)
            if '분석 요약' in xls.sheet_names:
                existing_summary = pd.read_excel(xls, sheet_name='분석 요약')
            if 'History' in xls.sheet_names:
                existing_history = pd.read_excel(xls, sheet_name='History')
            st.success("업로드된 엑셀을 읽었습니다. 이 파일에 결과를 이어 붙여서 새 파일을 내려드립니다.")
        except Exception as e:
            st.warning(f"엑셀을 읽는 중 문제가 발생했습니다: {e}")

# ----------------------
# Inputs
# ----------------------
with st.form("inputs"):
    c1, c2, c3 = st.columns(3)
    site_name = c1.text_input("0. 현장명", value="미지정")
    num_chargers = c2.number_input("5. 설치할 총 충전기 대수", min_value=1, value=10, step=1)
    usage_days   = c3.number_input("7. 정산기간 사용일수(일)", min_value=1, value=30, step=1)

    st.markdown("**원가 입력**")
    cc1, cc2, cc3 = st.columns(3)
    cost_install_ops  = cc1.number_input("1-1. 1기당 영업+시공비(원)", min_value=0.0, value=600_000.0, step=10_000.0)
    cost_hw_materials = cc2.number_input("1-2. 1기당 충전기+부자재(원)", min_value=0.0, value=700_000.0, step=10_000.0)
    per_unit_gross_cost = cost_install_ops + cost_hw_materials
    cc3.metric("기당 총 원가", f"{per_unit_gross_cost:,.0f} 원")

    st.markdown("**보조금 설정**")
    tiered = st.checkbox("계단식 보조금 자동 적용 (1기 2,200,000 / 2–5기 2,000,000 / 6기+ 1,800,000)", value=True)
    manual_subsidy_per_unit = st.number_input("2*. (수동모드 전용) 1기당 보조금(원)", min_value=0.0, value=0.0, step=10_000.0, disabled=tiered)

    st.markdown("**상수/스펙 (수동 체크 시 값 사용)**")
    sc1, sc2, sc3, sc4 = st.columns(4)
    m_cap = sc1.checkbox("용량 수동", value=False)
    capacity_kw = sc1.number_input("충전기 용량 (kW/기)", min_value=0.1, value=7.0, step=0.1, disabled=not m_cap)

    m_base = sc2.checkbox("기본요금 수동", value=False)
    base_fee_per_kw = sc2.number_input("전력 기본요금 (원/kW·월)", min_value=0.0, value=2390.0, step=10.0, disabled=not m_base)

    m_comms = sc3.checkbox("통신비 수동", value=False)
    comms_per_charger = sc3.number_input("통신비 (원/기·월)", min_value=0.0, value=5500.0, step=100.0, disabled=not m_comms)

    m_mgmt = sc4.checkbox("관리비 수동", value=False)
    mgmt_per_charger = sc4.number_input("관리비 (원/기·월)", min_value=0.0, value=5000.0, step=100.0, disabled=not m_mgmt)

    st.markdown("**총액 수동(활성 시 1-1/1-2/1기당 보조금은 무시)**")
    t1, t2 = st.columns(2)
    manual_total_invest_on  = t1.checkbox("총 투자비 수동 (보조금 차감 후)", value=False)
    manual_total_invest     = t1.number_input("총 투자비(원, 보조금 차감 후)", min_value=0.0, value=0.0, step=100_000.0, disabled=not manual_total_invest_on)

    manual_total_subsidy_on = t2.checkbox("총 보조금 수동", value=False)
    manual_total_subsidy    = t2.number_input("총 보조금(원)", min_value=0.0, value=0.0, step=100_000.0, disabled=not manual_total_subsidy_on)

    st.markdown("**요금/사용량**")
    p1, p2, p3 = st.columns(3)
    charging_fee_kwh     = p1.number_input("3. 1kWh 당 고객 충전요금(원)", min_value=0.0, value=350.0, step=10.0)
    electricity_rate_kwh = p2.number_input("4. 1kWh 당 전력량요금(원)", min_value=0.0, value=140.0, step=10.0)
    period_total_kwh     = p3.number_input("6. 정산기간 총 사용량 (kWh)", min_value=0.0, value=1800.0, step=10.0)

    submitted = st.form_submit_button("분석 실행")

# ----------------------
# Compute
# ----------------------
if submitted:
    # constants
    DEFAULT_BASE = 2390.0
    DEFAULT_COMMS = 5500.0
    DEFAULT_MGMT = 5000.0
    DEFAULT_CAP = 7.0

    base_val  = base_fee_per_kw if m_base else DEFAULT_BASE
    comms_val = comms_per_charger if m_comms else DEFAULT_COMMS
    mgmt_val  = mgmt_per_charger if m_mgmt else DEFAULT_MGMT
    cap_val   = capacity_kw if m_cap else DEFAULT_CAP

    monthly_fixed_per_charger = cap_val * base_val + comms_val + mgmt_val
    daily_fixed_per_charger = monthly_fixed_per_charger / 30.0

    # costs & subsidy
    total_gross_capex_auto = per_unit_gross_cost * num_chargers
    if tiered:
        total_subsidy_auto = calc_tiered_subsidy(num_chargers)
        avg_subsidy_auto = total_subsidy_auto / num_chargers
    else:
        total_subsidy_auto = manual_subsidy_per_unit * num_chargers
        avg_subsidy_auto = manual_subsidy_per_unit

    initial_investment_auto = total_gross_capex_auto - total_subsidy_auto

    # overrides
    if manual_total_invest_on:
        initial_investment = manual_total_invest
        if manual_total_subsidy_on:
            total_subsidy = manual_total_subsidy
        else:
            total_subsidy = max(total_gross_capex_auto - initial_investment, 0.0)
    else:
        if manual_total_subsidy_on:
            total_subsidy = manual_total_subsidy
            initial_investment = total_gross_capex_auto - total_subsidy
        else:
            total_subsidy = total_subsidy_auto
            initial_investment = initial_investment_auto

    grant_surplus = max(total_subsidy - total_gross_capex_auto, 0.0)
    initial_investment = max(initial_investment, 0.0)
    investment_per_charger = (initial_investment / num_chargers) if num_chargers > 0 else 0.0
    avg_subsidy_per_unit = (total_subsidy / num_chargers) if num_chargers > 0 else 0.0

    # usage → economics
    total_capacity_kw = cap_val * num_chargers
    invest_per_kw = (initial_investment / total_capacity_kw) if total_capacity_kw > 0 else 0.0

    daily_total_kwh = period_total_kwh / usage_days
    daily_kwh_per_charger = daily_total_kwh / num_chargers

    daily_revenue_per_charger = daily_kwh_per_charger * charging_fee_kwh
    daily_energy_cost_per_charger = daily_kwh_per_charger * electricity_rate_kwh
    daily_profit_per_charger = daily_revenue_per_charger - daily_energy_cost_per_charger - daily_fixed_per_charger

    revenue_30d_per_charger = daily_revenue_per_charger * 30
    energy_cost_30d_per_charger = daily_energy_cost_per_charger * 30
    profit_30d_per_charger = revenue_30d_per_charger - energy_cost_30d_per_charger - monthly_fixed_per_charger

    revenue_1y_per_charger = daily_revenue_per_charger * 365
    energy_cost_1y_per_charger = daily_energy_cost_per_charger * 365
    fixed_1y_per_charger = monthly_fixed_per_charger * 12
    profit_1y_per_charger = revenue_1y_per_charger - energy_cost_1y_per_charger - fixed_1y_per_charger

    monthly_revenue_total = revenue_30d_per_charger * num_chargers
    monthly_cost_total = energy_cost_30d_per_charger * num_chargers + monthly_fixed_per_charger * num_chargers
    monthly_profit_total = monthly_revenue_total - monthly_cost_total
    monthly_profit_per_kw = (monthly_profit_total / total_capacity_kw) if total_capacity_kw > 0 else 0.0

    subsidy_coverage_ratio = (total_subsidy / total_gross_capex_auto) if total_gross_capex_auto > 0 else None
    gross_payback_months = (total_gross_capex_auto / monthly_profit_total) if monthly_profit_total > 0 else float('inf')

    if initial_investment > 0 and monthly_profit_total > 0:
        payback_period_months = initial_investment / monthly_profit_total
    elif initial_investment == 0:
        payback_period_months = 0.0
    else:
        payback_period_months = float('inf')

    roi_per_charger = (profit_1y_per_charger / investment_per_charger * 100.0) if investment_per_charger > 0 else None

    # ----------------------
    # Display
    # ----------------------
    if grant_surplus > 0:
        st.subheader("분석 결과 (간단 모드)")
        st.info("보조금이 총원가를 초과하여 초기 투자금이 없습니다.")
        a1, a2 = st.columns(2)
        a1.metric("1기당 연간 순이익", f"{profit_1y_per_charger:,.0f} 원/년")
        a2.metric("총 연간 순이익", f"{profit_1y_per_charger * num_chargers:,.0f} 원/년")
        b1, b2, b3 = st.columns(3)
        b1.metric("월 평균 순이익(총합)", f"{monthly_profit_total:,.0f} 원/월")
        b2.metric("보조금 잉여(초기 유입)", f"{grant_surplus:,.0f} 원")
        b3.metric("보조금 커버리지", "N/A" if subsidy_coverage_ratio is None else f"{subsidy_coverage_ratio:.2f} x")
    else:
        st.subheader("분석 결과")
        m1, m2, m3 = st.columns(3)
        m1.metric("1기당 1일 순이익", f"{daily_profit_per_charger:,.0f} 원")
        m2.metric("1기당 30일 순이익", f"{profit_30d_per_charger:,.0f} 원")
        m3.metric("1기당 연간 순이익", f"{profit_1y_per_charger:,.0f} 원")
        mm1, mm2, mm3 = st.columns(3)
        mm1.metric("총 월 순이익", f"{monthly_profit_total:,.0f} 원/월")
        mm2.metric("회수기간", "회수 불가" if payback_period_months == float('inf') else f"{payback_period_months:.1f} 개월")
        mm3.metric("ROI(연간)", "N/A" if roi_per_charger is None else f"{roi_per_charger:.2f} %")

    # ----------------------
    # Excel export
    # ----------------------
    now_label = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    col_name = f"{site_name} [{now_label}]"

    df_two_col = build_summary_two_col(
        site_name, num_chargers,
        cost_install_ops, cost_hw_materials, per_unit_gross_cost,
        investment_per_charger, initial_investment,
        monthly_revenue_total, monthly_cost_total, monthly_profit_total,
        profit_1y_per_charger, payback_period_months, roi_per_charger,
        total_subsidy, avg_subsidy_per_unit, total_gross_capex_auto,
        tiered,
        base_val, comms_val, mgmt_val, cap_val,
        m_base, m_comms, m_mgmt, m_cap,
        total_capacity_kw, invest_per_kw, monthly_profit_per_kw,
        manual_total_invest_on, manual_total_invest,
        manual_total_subsidy_on, manual_total_subsidy,
        grant_surplus, subsidy_coverage_ratio, gross_payback_months
    )

    wide_summary = append_summary_as_column(existing_summary, df_two_col, col_name)

    # History row
    history_row = {
        'Run ID': now_label,
        '현장명': site_name,
        '설치 대수': num_chargers,
        '1-1 영업+시공(기당 원)': int(cost_install_ops),
        '1-2 충전기+부자재(기당 원)': int(cost_hw_materials),
        '기당 총 원가(원)': int(per_unit_gross_cost),
        '총 원가(원)': int(total_gross_capex_auto),
        '총 보조금(원)': int(total_subsidy),
        '총 보조금(수동 사용?)': manual_total_subsidy_on,
        '총 보조금(수동 입력값)': int(manual_total_subsidy) if manual_total_subsidy_on else None,
        '총 투자비(원, 보조금 차감)': int(initial_investment),
        '총 투자비(수동 사용?)': manual_total_invest_on,
        '총 투자비(수동 입력값)': int(manual_total_invest) if manual_total_invest_on else None,
        '기기당 실투자비(원)': int(investment_per_charger),
        '충전기 용량(기당 kW)': float(cap_val),
        '총 충전기 용량(kW)': float(total_capacity_kw),
        '투자금/총 용량(원/kW)': float(invest_per_kw),
        '월 순이익/총 용량(원/kW·월)': float(monthly_profit_per_kw),
        '월 총 매출(원)': int(monthly_revenue_total),
        '월 총 비용(원)': int(monthly_cost_total),
        '월 총 순이익(원)': int(monthly_profit_total),
        '기당 연 순이익(원)': int(profit_1y_per_charger),
        '투자비 회수 기간(개월)': None if payback_period_months == float('inf') else round(payback_period_months, 2),
        '연간 ROI(%)': None if roi_per_charger is None else float(roi_per_charger),
        '전력 기본요금(kW·월)': int(base_val),
        '통신비(기·월)': int(comms_val),
        '관리비(기·월)': int(mgmt_val),
        '보조금 잉여(초기 유입 원)': int(grant_surplus),
        '보조금 커버리지(보조금/총원가)': None if subsidy_coverage_ratio is None else float(subsidy_coverage_ratio),
        '총원가 기준 회수기간(개월)': None if gross_payback_months == float('inf') else round(gross_payback_months, 2),
        '보조금 방식(1기 기준)': ("계단식" if tiered else "1기당 수동"),
    }
    history_df_new = pd.DataFrame([history_row])
    history_df = pd.concat([existing_history, history_df_new], ignore_index=True) if existing_history is not None else history_df_new

    # Build Excel in memory
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        wide_summary.to_excel(writer, index=False, sheet_name='분석 요약')
        history_df.to_excel(writer, index=False, sheet_name='History')
    st.download_button(
        label="⬇️ 엑셀 다운로드 (분석 요약 + History)",
        data=buffer.getvalue(),
        file_name=f"EV_Charger_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.success("완료! 다운로드 버튼으로 결과 파일을 저장하세요. 기존 마스터를 올렸다면 이어붙여서 내려갑니다.")
