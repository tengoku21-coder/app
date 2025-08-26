# -*- coding: utf-8 -*-
"""
EV Charger Business Analyzer (Web / Streamlit) — Fixed & Enhanced
v2025-08-26

변경 요약
- '1년차 평균 월 순이익(가중)' 라벨 정정(화면 표시)
- ROI 2종 병행: (1) 순투자 기준(보조금 차감), (2) 총원가 기준(추가 지표/엑셀 기록)
- 과대표시 방지 경고: 보조금 커버리지 ≥1.0, 순투자 과소
- 이용률(이론상 최대 대비) 표시 및 고이용률 경고
- 엑셀: '1년차 총 순이익(총합)', 'ROI(총원가 기준, 1년차 가중/정상)' 등 추가 기록
"""
import io
from datetime import datetime
from pathlib import Path
from base64 import b64encode

import pandas as pd
import streamlit as st

# ======================
# 상단 로고/헤더 (잘림 방지, 가로 정렬)
# ======================
APP_DIR = Path(__file__).parent
LOGO_FILE = "ts-logo.png"      # 로고 파일명(선택)
LOGO_HEIGHT = 48               # 로고 세로 높이(px)

# 페이지 설정(가장 먼저)
logo_path = None
for cand in [APP_DIR / LOGO_FILE, APP_DIR / "assets" / LOGO_FILE]:
    if cand.exists():
        logo_path = str(cand)
        break

st.set_page_config(
    page_title="EV Charger Analyzer (Web)",
    page_icon=logo_path if logo_path else None,
    layout="wide",
)

# 스타일: 상단 패딩 넉넉히(잘림 방지), 로고/제목 한 줄 가로 정렬
st.markdown(
    f"""
    <style>
      .block-container {{ padding-top: 2.2rem !important; }}
      .tsct-header {{ display:flex; align-items:center; gap:16px; margin:0 0 .5rem 0; }}
      .tsct-logo {{ height:{LOGO_HEIGHT}px; width:auto; display:block; }}
      .tsct-title {{ font-size:1.6rem; font-weight:700; margin:0; }}
      .tsct-sub   {{ margin:.2rem 0 0 0; color:#6c757d; font-size:.95rem; }}
      @media (max-width: 600px) {{
        .tsct-title {{ font-size:1.25rem; }}
        .tsct-logo  {{ height:{max(36, LOGO_HEIGHT-12)}px; }}
      }}
    </style>
    """,
    unsafe_allow_html=True
)

def _img_b64(p: str) -> str:
    with open(p, "rb") as f:
        return b64encode(f.read()).decode()

logo_b64 = _img_b64(logo_path) if logo_path else None
st.markdown(
    f"""
    <div class="tsct-header">
      {'<img class="tsct-logo" src="data:image/png;base64,' + logo_b64 + '"/>' if logo_b64 else ''}
      <div>
        <h2 class="tsct-title">⚡ 전기차 충전기 사업성 분석기 — <span style="font-weight:600">Web</span></h2>
        <p class="tsct-sub">데스크탑 계산 로직을 그대로 웹에서 사용</p>
      </div>
    </div>
    """,
    unsafe_allow_html=True
)

# ======================
# 유틸 함수
# ======================
def calc_tiered_subsidy(n: int) -> int:
    """총 보조금: 1기=2,200,000 / 2~5기=2,000,000 / 6기+=1,800,000"""
    if n <= 0:
        return 0
    total = 2_200_000 if n >= 1 else 0
    if n >= 2:
        total += min(n - 1, 4) * 2_000_000
    if n >= 6:
        total += (n - 5) * 1_800_000
    return total

def append_summary_as_column(existing_wide, df_two_col, col_name: str) -> pd.DataFrame:
    """분석 요약(열 누적): 각 실행을 새 열로 추가"""
    if existing_wide is None or '항목' not in getattr(existing_wide, "columns", []):
        return df_two_col.rename(columns={'값': col_name}).copy()
    wide = existing_wide.copy()
    merged = pd.merge(wide, df_two_col, on='항목', how='outer').rename(columns={'값': col_name})
    order_map = {v: i for i, v in enumerate(df_two_col['항목'].tolist())}
    merged['__ord__'] = merged['항목'].map(order_map).fillna(9999).astype(int)
    return merged.sort_values('__ord__').drop(columns='__ord__')

def build_summary_two_col(
    site_name, num_chargers,
    cost_install_ops, cost_hw_materials, per_unit_gross_cost,
    investment_per_charger, initial_investment,
    # 1년차(프로모션 가중)
    monthly_revenue_blended_total, monthly_cost_total, monthly_profit_blended_total,
    profit_1y_per_charger_blended, payback_period_months, roi_per_charger_blended,
    # 장기(정상요금 기준)
    monthly_profit_steady_total,  profit_1y_per_charger_steady, roi_per_charger_steady,
    # 원가/보조금/상수
    total_subsidy, avg_subsidy_per_unit, total_gross_capex, is_tiered,
    base_fee_per_kw, comms_per_charger, mgmt_per_charger, capacity_kw,
    m_base, m_comms, m_mgmt, m_cap,
    total_capacity_kw, invest_per_kw, monthly_profit_per_kw_blended,
    m_total_inv_on, m_total_inv_val, m_total_sub_on, m_total_sub_val,
    grant_surplus, subsidy_coverage_ratio, gross_payback_months,
    # 표시용
    monthly_profit_regular_total, monthly_profit_promo_total,
    promo_on, promo_months, promo_price_kwh, normal_price_kwh
) -> pd.DataFrame:
    """항목/값 2열 요약 (요약 시트에 열로 붙일 원본)"""
    const_mode = f"{'수동' if m_base else '자동'}/{'수동' if m_comms else '자동'}/{'수동' if m_mgmt else '자동'}/{'수동' if m_cap else '자동'}"
    subsidy_mode_label = "계단식" if is_tiered else "1기당 수동"
    items = [
        '현장명','총 설치 대수',
        '1-1 영업+시공(기당)','1-2 충전기+부자재(기당)','기당 총 원가',
        '총 원가(설치비 합계)','총 보조금(최종)','총 투자비(최종, 보조금 차감)',
        '총 보조금(수동 입력값)','총 투자비(수동 입력값)',
        '기당 평균 보조금','기기당 실투자비','보조금 방식(1기 기준)','---',
        '충전기 용량(기당 kW)','총 충전기 용량(kW)',
        '전력 기본요금(kW·월)','통신비(기·월)','관리비(기·월)','상수/스펙 입력 모드','---',
        '정상가(원/kWh)','프로모션가(원/kWh)','프로모션 개월','---',
        # 1년차(가중)
        '월 총 순이익(정상가, 총합)','월 총 순이익(프로모션가, 총합)','월 총 순이익(가중, 총합)',
        '월 총 매출(가중, 총합)','월 총 비용(총합)','월 순이익/총 용량(원/kW·월)','기기당 연간 순이익(1년차 가중)','ROI(연간, 1년차 가중)','회수기간(개월, 동적)','---',
        # 장기(정상요금)
        '월 총 순이익(포스트 프로모션, 총합)','기기당 연간 순이익(포스트 프로모션)','ROI(연간, 포스트 프로모션)','---',
        '보조금 잉여(초기 유입 원)','보조금 커버리지(보조금/총원가)','총원가 기준 회수기간(개월)'
    ]
    vals = [
        site_name, f"{num_chargers} 대",
        f"{cost_install_ops:,.0f} 원", f"{cost_hw_materials:,.0f} 원", f"{per_unit_gross_cost:,.0f} 원",
        f"{total_gross_capex:,.0f} 원", f"{total_subsidy:,.0f} 원", f"{initial_investment:,.0f} 원",
        ("" if not m_total_sub_on else f"{m_total_sub_val:,.0f} 원"),
        ("" if not m_total_inv_on else f"{m_total_inv_val:,.0f} 원"),
        f"{avg_subsidy_per_unit:,.0f} 원", f"{investment_per_charger:,.0f} 원", subsidy_mode_label, "",
        f"{capacity_kw:,.2f} kW", f"{total_capacity_kw:,.2f} kW",
        f"{base_fee_per_kw:,.0f} 원", f"{comms_per_charger:,.0f} 원", f"{mgmt_per_charger:,.0f} 원", const_mode, "",
        f"{normal_price_kwh:,.0f}", (f"{promo_price_kwh:,.0f}" if promo_on else "—"), (f"{promo_months}" if promo_on else "0"), "",
        f"{monthly_profit_regular_total:,.0f} 원", f"{monthly_profit_promo_total:,.0f} 원",
        f"{monthly_profit_blended_total:,.0f} 원",
        f"{monthly_revenue_blended_total:,.0f} 원", f"{monthly_cost_total:,.0f} 원",
        f"{monthly_profit_per_kw_blended:,.0f} 원/kW·월", f"{profit_1y_per_charger_blended:,.0f} 원",
        ("N/A" if roi_per_charger_blended is None else f"{roi_per_charger_blended:.2f} %"),
        ("회수 불가" if payback_period_months == float('inf') else f"{payback_period_months:.1f} 개월"), "",
        f"{monthly_profit_steady_total:,.0f} 원", f"{profit_1y_per_charger_steady:,.0f} 원",
        ("N/A" if roi_per_charger_steady is None else f"{roi_per_charger_steady:.2f} %"), "",
        f"{grant_surplus:,.0f} 원",
        ("N/A" if subsidy_coverage_ratio is None else f"{subsidy_coverage_ratio:.2f} x"),
        ("∞" if gross_payback_months == float('inf') else f"{gross_payback_months:.1f}")
    ]
    return pd.DataFrame({'항목': items, '값': vals})

# ======================
# (선택) 기존 마스터 엑셀 업로드
# ======================
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

# ======================
# Inputs (LIVE)
# ======================
c1, c2, c3 = st.columns(3)
site_name = c1.text_input("0. 현장명", value="미지정")
num_chargers = c2.number_input("5. 설치할 총 충전기 대수", min_value=1, value=10, step=1)
usage_days   = c3.number_input("7. 정산기간 사용일수(일)", min_value=1, value=30, step=1)

st.markdown("**총액 수동(활성 시 1-1/1-2/1기당 보조금은 비활성화)**")
t1, t2 = st.columns(2)
manual_total_invest_on  = t1.checkbox("총 투자비 수동 (보조금 차감 후)", value=False)
manual_total_invest     = t1.number_input("총 투자비(원, 보조금 차감 후)", min_value=0.0, value=0.0,
                                          step=100_000.0, disabled=not manual_total_invest_on)
manual_total_subsidy_on = t2.checkbox("총 보조금 수동", value=False)
manual_total_subsidy    = t2.number_input("총 보조금(원)", min_value=0.0, value=0.0,
                                          step=100_000.0, disabled=not manual_total_subsidy_on)
disable_costs = manual_total_invest_on or manual_total_subsidy_on

st.markdown("**원가 입력**")
cc1, cc2, cc3 = st.columns(3)
cost_install_ops  = cc1.number_input("1-1. 1기당 영업+시공비(원)", min_value=0.0, value=600_000.0,
                                     step=10_000.0, disabled=disable_costs)
cost_hw_materials = cc2.number_input("1-2. 1기당 충전기+부자재(원)", min_value=0.0, value=700_000.0,
                                     step=10_000.0, disabled=disable_costs)
per_unit_gross_cost = cost_install_ops + cost_hw_materials
cc3.metric("기당 총 원가", f"{per_unit_gross_cost:,.0f} 원")

st.markdown("**보조금 설정**")
tiered = st.checkbox("계단식 보조금 자동 적용 (1기 2,200,000 / 2–5기 2,000,000 / 6기+ 1,800,000)",
                     value=True, disabled=manual_total_subsidy_on)
manual_subsidy_per_unit = st.number_input("2*. (수동모드 전용) 1기당 보조금(원)",
                     min_value=0.0, value=0.0, step=10_000.0,
                     disabled=(tiered or manual_total_subsidy_on))

st.markdown("**상수/스펙 (수동 체크 시 값 사용)**")
sc1, sc2, sc3, sc4 = st.columns(4)
m_cap  = sc1.checkbox("용량 수동", value=False)
cap_val_input = sc1.number_input("충전기 용량 (kW/기)", min_value=0.1, value=7.0, step=0.1, disabled=not m_cap)
m_base = sc2.checkbox("기본요금 수동", value=False)
base_fee_input = sc2.number_input("전력 기본요금 (원/kW·월)", min_value=0.0, value=2390.0, step=10.0, disabled=not m_base)
m_comms = sc3.checkbox("통신비 수동", value=False)
comms_input = sc3.number_input("통신비 (원/기·월)", min_value=0.0, value=5500.0, step=100.0, disabled=not m_comms)
m_mgmt = sc4.checkbox("관리비 수동", value=False)
mgmt_input = sc4.number_input("관리비 (원/기·월)", min_value=0.0, value=5000.0, step=100.0, disabled=not m_mgmt)

st.markdown("**요금/사용량**")
p1, p2, p3 = st.columns(3)
charging_fee_kwh     = p1.number_input("3. 1kWh 당 고객 충전요금(원)", min_value=0.0, value=350.0, step=10.0)
electricity_rate_kwh = p2.number_input("4. 1kWh 당 전력량요금(원)", min_value=0.0, value=140.0, step=10.0)
period_total_kwh     = p3.number_input("6. 정산기간 총 사용량 (kWh)", min_value=0.0, value=1800.0, step=10.0)

# 총 사용량 단위(정산기간 총량 / 월 총량)
usage_basis = st.radio(
    "총 사용량 단위 선택",
    ["정산기간 총량 (아래 '정산기간 사용일수' 기준)", "월 총량 (해당 월 전체 kWh)"],
    index=0, horizontal=True
)

st.markdown("**프로모션(할인) 요금**")
pp1, pp2, pp3 = st.columns(3)
promo_on        = pp1.checkbox("프로모션 사용", value=False, help="체크하면 아래 개월/단가가 반영됩니다.")
promo_months    = pp2.number_input("프로모션 기간(개월)", min_value=0, max_value=12, value=6, step=1, disabled=not promo_on)
promo_price_kwh = pp3.number_input("프로모션 요금 (원/kWh)", min_value=0.0, value=168.0, step=1.0, disabled=not promo_on)

# 제출 버튼(폼 아님)
submitted = st.button("분석 실행")

# ======================
# Compute & Display
# ======================
if submitted:
    # 기본 상수
    DEFAULT_BASE, DEFAULT_COMMS, DEFAULT_MGMT, DEFAULT_CAP = 2390.0, 5500.0, 5000.0, 7.0
    base_val  = base_fee_input if m_base else DEFAULT_BASE
    comms_val = comms_input    if m_comms else DEFAULT_COMMS
    mgmt_val  = mgmt_input     if m_mgmt else DEFAULT_MGMT
    cap_val   = cap_val_input  if m_cap  else DEFAULT_CAP

    monthly_fixed_per_charger = cap_val * base_val + comms_val + mgmt_val
    daily_fixed_per_charger   = monthly_fixed_per_charger / 30.0

    # --- 1기당 월 kWh 계산 ---
    if usage_basis.startswith("월 총량"):
        # 입력된 period_total_kwh가 '모든 기기의 월 총량'
        monthly_kwh_pc = (period_total_kwh / max(num_chargers, 1))
    else:
        # 정산기간 총량 → 1일 평균 → 30일 환산 → 1기당
        daily_total_kwh = period_total_kwh / max(usage_days, 1)
        daily_kwh_per_charger = daily_total_kwh / max(num_chargers, 1)
        monthly_kwh_pc = daily_kwh_per_charger * 30

    # 이론상 최대(연속 24h 충전): cap[kW] * 24h * 30일
    theoretical_max_kwh_pc = cap_val * 24 * 30
    utilization = monthly_kwh_pc / theoretical_max_kwh_pc if theoretical_max_kwh_pc > 0 else 0.0
    if utilization > 1.0:
        st.warning(
            f"계산된 **1기당 월 사용량 {monthly_kwh_pc:,.0f} kWh**가 "
            f"이론상 최대치 **{theoretical_max_kwh_pc:,.0f} kWh**(24h×30일×{cap_val:g}kW)를 초과합니다. "
            "‘총 사용량 단위’/‘정산기간 사용일수’/‘설치 대수’ 입력을 다시 확인해 주세요."
        )

    # 에너지 비용(1기)
    monthly_energy_cost_pc = monthly_kwh_pc * electricity_rate_kwh

    # 정상가 & 프로모션 월 이익(1기)
    monthly_revenue_pc_regular = monthly_kwh_pc * charging_fee_kwh
    monthly_profit_pc_regular  = monthly_revenue_pc_regular - monthly_energy_cost_pc - monthly_fixed_per_charger

    monthly_revenue_pc_promo   = monthly_kwh_pc * promo_price_kwh
    monthly_profit_pc_promo    = monthly_revenue_pc_promo - monthly_energy_cost_pc - monthly_fixed_per_charger

    # Fleet totals
    monthly_profit_regular_total  = monthly_profit_pc_regular * num_chargers
    monthly_profit_promo_total    = monthly_profit_pc_promo   * num_chargers
    monthly_revenue_regular_total = monthly_revenue_pc_regular * num_chargers
    monthly_revenue_promo_total   = monthly_revenue_pc_promo   * num_chargers
    monthly_cost_total            = (monthly_energy_cost_pc + monthly_fixed_per_charger) * num_chargers  # 가격 무관 비용

    # 1년차(프로모션 가중)
    pm = promo_months if promo_on else 0
    pm = int(max(0, min(12, pm)))
    w_p, w_r = pm/12.0, 1 - (pm/12.0)
    monthly_profit_blended_total  = w_p * monthly_profit_promo_total  + w_r * monthly_profit_regular_total
    monthly_revenue_blended_total = w_p * monthly_revenue_promo_total + w_r * monthly_revenue_regular_total

    total_capacity_kw = cap_val * num_chargers
    monthly_profit_per_kw_blended = (monthly_profit_blended_total / total_capacity_kw) if total_capacity_kw > 0 else 0.0
    profit_1y_per_charger_blended = (pm * monthly_profit_pc_promo + (12 - pm) * monthly_profit_pc_regular)

    # 장기(포스트-프로모션=정상요금)
    monthly_profit_steady_total  = monthly_profit_regular_total
    profit_1y_per_charger_steady = monthly_profit_pc_regular * 12

    # 원가/보조금
    total_gross_capex_auto = (cost_install_ops + cost_hw_materials) * num_chargers
    if tiered:
        total_subsidy_auto = calc_tiered_subsidy(num_chargers); avg_subsidy_auto = total_subsidy_auto/num_chargers
    else:
        total_subsidy_auto = manual_subsidy_per_unit * num_chargers; avg_subsidy_auto = manual_subsidy_per_unit
    initial_investment_auto = total_gross_capex_auto - total_subsidy_auto

    # 총액 수동 오버라이드
    if manual_total_invest_on:
        initial_investment = manual_total_invest
        total_subsidy = (manual_total_subsidy if manual_total_subsidy_on
                         else max(total_gross_capex_auto - initial_investment, 0.0))
    else:
        if manual_total_subsidy_on:
            total_subsidy = manual_total_subsidy
            initial_investment = total_gross_capex_auto - total_subsidy
        else:
            total_subsidy = total_subsidy_auto
            initial_investment = initial_investment_auto

    # 파생 지표
    grant_surplus = max(total_subsidy - total_gross_capex_auto, 0.0)
    initial_investment = max(initial_investment, 0.0)
    investment_per_charger = (initial_investment / num_chargers) if num_chargers > 0 else 0.0
    avg_subsidy_per_unit   = (total_subsidy / num_chargers) if num_chargers > 0 else 0.0
    invest_per_kw          = (initial_investment / total_capacity_kw) if total_capacity_kw > 0 else 0.0

    # 회수기간: 월별 시뮬레이션(프로모션 pm개월 → 정상)
    HORIZON = 240  # 20년
    series = ([monthly_profit_promo_total] * pm) + ([monthly_profit_regular_total] * (HORIZON - pm))
    def months_to_recover(target: float, seq) -> float:
        if target <= 0:
            return 0.0
        s = 0.0
        for i, m in enumerate(seq, start=1):
            s += m
            if s >= target:
                return float(i)
        return float('inf')

    subsidy_coverage_ratio = (total_subsidy / total_gross_capex_auto) if total_gross_capex_auto > 0 else None
    gross_payback_months   = months_to_recover(total_gross_capex_auto, series)
    payback_period_months  = months_to_recover(initial_investment, series)

    # ROI: 1년차(가중) / 장기(정상) — 순투자 기준
    roi_per_charger_blended = (profit_1y_per_charger_blended / investment_per_charger * 100.0) if investment_per_charger > 0 else None
    roi_per_charger_steady  = (profit_1y_per_charger_steady  / investment_per_charger * 100.0) if investment_per_charger > 0 else None

    # === 추가 지표: 총원가 기준 ROI + 1년차 총 순이익 ===
    profit_1y_total_blended = profit_1y_per_charger_blended * num_chargers
    roi_per_charger_blended_gross = (
        profit_1y_per_charger_blended / per_unit_gross_cost * 100.0
    ) if per_unit_gross_cost > 0 else None
    roi_per_charger_steady_gross = (
        profit_1y_per_charger_steady / per_unit_gross_cost * 100.0
    ) if per_unit_gross_cost > 0 else None

    # ----- 표시 -----
    st.subheader("분석 결과")
    cA, cB, cC = st.columns(3)
    cA.metric("1기당 1일 순이익(정상가)", f"{(monthly_profit_pc_regular/30):,.0f} 원/일")
    cB.metric("1기당 30일 순이익(정상가)", f"{monthly_profit_pc_regular:,.0f} 원/월")
    cC.metric("회수기간(동적)", "회수 불가" if payback_period_months == float('inf') else f"{payback_period_months:.1f} 개월")

    # 이용률 정보/경고
    st.caption(
        f"평균 이용률(1기) ≈ {utilization*100:.1f}% · "
        f"이론상 최대 {theoretical_max_kwh_pc:,.0f} kWh/월 대비"
    )
    if utilization >= 0.60:
        st.warning("이용률이 60% 이상입니다. 사용량/대수/일수 입력값을 다시 확인해 보세요.")

    if promo_on:
        st.caption(f"프로모션 기준 1기당 월 순이익: {monthly_profit_pc_promo:,.0f} 원/월 · 적용 {pm}개월")

    st.markdown("**① 1년차(프로모션 반영)**")
    y1_1, y1_2, y1_3 = st.columns(3)
    # 라벨 정정: 평균 월 순이익(가중)
    y1_1.metric("1년차 평균 월 순이익(가중)", f"{monthly_profit_blended_total:,.0f} 원/월")
    y1_2.metric("1기당 연간 순이익(가중)", f"{profit_1y_per_charger_blended:,.0f} 원/년")
    y1_3.metric("ROI(연간, 가중·순투자)", "N/A" if roi_per_charger_blended is None else f"{roi_per_charger_blended:.2f} %")

    # 총원가 기준 ROI 안내 캡션(참고치)
    st.caption(
        "참고: ROI(총원가 기준, 1년차 가중) = "
        + ("N/A" if roi_per_charger_blended_gross is None else f"{roi_per_charger_blended_gross:.2f} %")
        + f" · 1년차 총 순이익(총합) = {profit_1y_total_blended:,.0f} 원/년"
    )

    st.markdown("**② 포스트-프로모션(정상요금) 기준**")
    ss_1, ss_2, ss_3 = st.columns(3)
    ss_1.metric("총 월 순이익(정상요금)", f"{monthly_profit_steady_total:,.0f} 원/월")
    ss_2.metric("1기당 연간 순이익(정상요금)", f"{profit_1y_per_charger_steady:,.0f} 원/년")
    ss_3.metric("ROI(연간, 정상요금·순투자)", "N/A" if roi_per_charger_steady is None else f"{roi_per_charger_steady:.2f} %")

    # 보조금/순투자 과대표시 경고
    if subsidy_coverage_ratio is not None:
        if subsidy_coverage_ratio >= 1.0:
            st.warning(
                "보조금이 총원가를 초과(커버리지 ≥ 1.0)합니다. "
                "‘순투자 기준 ROI’는 의미가 없고, ‘총원가 기준 ROI’를 참고하세요."
            )
        elif investment_per_charger < max(200_000, per_unit_gross_cost * 0.15):
            st.info(
                "순투자(보조금 차감)가 매우 작아 ROI가 과대해 보일 수 있습니다. "
                "총원가 기준 ROI도 함께 확인하세요."
            )

    # ----- 엑셀 다운로드 (분석 요약 + History) -----
    now_label = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    col_name  = f"{site_name} [{now_label}]"

    df_two_col = build_summary_two_col(
        site_name, num_chargers,
        cost_install_ops, cost_hw_materials, per_unit_gross_cost,
        investment_per_charger, initial_investment,
        monthly_revenue_blended_total, monthly_cost_total, monthly_profit_blended_total,
        profit_1y_per_charger_blended, payback_period_months, roi_per_charger_blended,
        monthly_profit_steady_total,  profit_1y_per_charger_steady, roi_per_charger_steady,
        total_subsidy, avg_subsidy_per_unit, total_gross_capex_auto, tiered,
        base_val, comms_val, mgmt_val, cap_val,
        m_base, m_comms, m_mgmt, m_cap,
        total_capacity_kw, invest_per_kw, monthly_profit_per_kw_blended,
        manual_total_invest_on, manual_total_invest, manual_total_subsidy_on, manual_total_subsidy,
        grant_surplus, subsidy_coverage_ratio, gross_payback_months,
        monthly_profit_regular_total, monthly_profit_promo_total,
        promo_on, pm, promo_price_kwh, charging_fee_kwh
    )

    # 요약 시트에 추가 지표 붙이기(총원가 ROI/총 순이익)
    extras = pd.DataFrame({
        '항목': [
            '1년차 총 순이익(총합)',
            'ROI(총원가 기준, 1년차 가중)',
            'ROI(총원가 기준, 포스트 프로모션)'
        ],
        '값': [
            f"{profit_1y_total_blended:,.0f} 원",
            ("N/A" if roi_per_charger_blended_gross is None else f"{roi_per_charger_blended_gross:.2f} %"),
            ("N/A" if roi_per_charger_steady_gross  is None else f"{roi_per_charger_steady_gross:.2f} %"),
        ]
    })
    df_two_col = pd.concat([df_two_col, extras], ignore_index=True)

    wide_summary = append_summary_as_column(existing_summary, df_two_col, col_name)

    history_row = {
        'Run ID': now_label, '현장명': site_name, '설치 대수': num_chargers,
        # 요금/프로모션
        '정상가(원/kWh)': float(charging_fee_kwh), '프로모션 사용?': bool(promo_on),
        '프로모션 개월': int(pm), '프로모션가(원/kWh)': float(promo_price_kwh if promo_on else 0),
        # 월/연 이익(1년차/정상)
        '월 순이익(정상가, 총합)': float(monthly_profit_regular_total),
        '월 순이익(프로모션가, 총합)': float(monthly_profit_promo_total),
        '월 순이익(가중, 총합)': float(monthly_profit_blended_total),
        '월 순이익(정상요금, 총합/포스트)': float(monthly_profit_steady_total),
        '월 총 매출(가중, 총합)': float(monthly_revenue_blended_total),
        '월 총 비용(총합)': float(monthly_cost_total),
        '월 순이익/총 용량(원/kW·월)': float(monthly_profit_per_kw_blended),
        '연간 순이익(기당, 1년차 가중)': float(profit_1y_per_charger_blended),
        '연간 순이익(기당, 정상요금)': float(profit_1y_per_charger_steady),
        '1년차 총 순이익(총합)': float(profit_1y_total_blended),
        '회수기간(개월, 동적)': None if payback_period_months == float('inf') else round(payback_period_months, 2),
        'ROI(연간, 1년차 가중, %·순투자)': None if roi_per_charger_blended is None else float(roi_per_charger_blended),
        'ROI(연간, 정상요금, %·순투자)': None if roi_per_charger_steady  is None else float(roi_per_charger_steady),
        'ROI(총원가 기준, 1년차 가중, %)': None if roi_per_charger_blended_gross is None else float(roi_per_charger_blended_gross),
        'ROI(총원가 기준, 정상요금, %)': None if roi_per_charger_steady_gross  is None else float(roi_per_charger_steady_gross),
        # 원가/보조금/상수
        '1-1 영업+시공(기당 원)': int(cost_install_ops), '1-2 충전기+부자재(기당 원)': int(cost_hw_materials),
        '기당 총 원가(원)': int(per_unit_gross_cost), '총 원가(원)': int(total_gross_capex_auto),
        '총 보조금(원)': int(total_subsidy), '총 보조금(수동 사용?)': bool(manual_total_subsidy_on),
        '총 보조금(수동 입력값)': int(manual_total_subsidy) if manual_total_subsidy_on else None,
        '총 투자비(원, 보조금 차감)': int(initial_investment), '총 투자비(수동 사용?)': bool(manual_total_invest_on),
        '총 투자비(수동 입력값)': int(manual_total_invest) if manual_total_invest_on else None,
        '기기당 실투자비(원)': int(investment_per_charger),
        '충전기 용량(기당 kW)': float(cap_val), '총 충전기 용량(kW)': float(total_capacity_kw),
        '투자금/총 용량(원/kW)': float(invest_per_kw),
        '전력 기본요금(kW·월)': int(base_val), '통신비(기·월)': int(comms_val), '관리비(기·월)': int(mgmt_val),
        '보조금 잉여(초기 유입 원)': int(grant_surplus),
        '보조금 커버리지(보조금/총원가)': None if subsidy_coverage_ratio is None else float(subsidy_coverage_ratio),
        '총원가 기준 회수기간(개월)': None if gross_payback_months == float('inf') else round(gross_payback_months, 2),
        '보조금 방식(1기 기준)': ("계단식" if tiered else "1기당 수동"),
        '평균 이용률(1기, %)': float(utilization*100.0),
    }
    history_df_new = pd.DataFrame([history_row])
    history_df = (pd.concat([existing_history, history_df_new], ignore_index=True)
                  if existing_history is not None else history_df_new)

    # --- 엑셀을 메모리에 생성 + 다운로드 버튼 ---
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        wide_summary.to_excel(writer, index=False, sheet_name='분석 요약')
        history_df.to_excel(writer, index=False, sheet_name='History')
    buffer.seek(0)

    st.download_button(
        label="⬇️ 엑셀 다운로드 (분석 요약 + History)",
        data=buffer.getvalue(),
        file_name=f"EV_Charger_Analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("좌측 입력 후 ‘분석 실행’을 누르면 결과와 엑셀 다운로드 버튼이 표시됩니다.")
