# -*- coding: utf-8 -*-
"""
IFRS17_Insurance-Liability-Model 계산 프로그램
"""

import xlwings as xw
import pandas as pd
import numpy as np
import re
import math
import os
from typing import Optional, Dict, Tuple, List

# 레이아웃 자동 설정 함수 import
try:
    from layout_setting import sync_layout
except ImportError:
    print("⚠ layout_setting.py를 찾을 수 없습니다. 레이아웃 자동 설정 기능이 비활성화됩니다.")
    sync_layout = None

# ==================== 전역 변수 ====================
wb = None
inp = None
da = None
cal = None
out = None

# 입력 레이블
LABELS = {
    "issue_age": "피보험자 나이(세)",
    "policy_term": "보험기간(년)",
    "annual_premium": "연납보험료",
    "sum_assured": "보험가입금액",
    "discount_rate": "할인율",
    "expense_ratio": "사업비율",
    "direct_expense": "직접사업비",
    "surrender_value": "해약환급금",
    "mortality_rate": "사망률",
}

LABELS_TWO_YEARS = {
    "premium_actual": "실제 보험료",
    "surrender_actual": "실제 해약환급금",
}

OUTPUT_LABELS = {
    "pv_inflow": "PV(유입)_보험료",
    "pv_outflow": "PV(유출)_보험금",
    "pv_surrender": "PV(유출)_해약환급금",
    "pv_expense": "PV(유출)_직접사업비",
    "bel": "최선추정부채(BEL)",
    "ra": "위험조정(RA)",
    "fcf": "이행현금흐름(FCF)",
    "csm": "보험계약마진(CSM)",
    "liability": "보험계약부채",
}


# ==================== 유틸리티 함수 ====================
def normalize(text):
    """공백과 줄바꿈 제거"""
    return re.sub(r"\s+", "", str(text or ""))


def get_input_value(sheet, label: str, label_col: str = "B", value_col: str = "C", max_rows: int = 200):
    """입력 시트에서 레이블(B열)로 값(C열) 찾기"""
    rng_labels = sheet[f"{label_col}1:{label_col}{max_rows}"]
    for cell in rng_labels:
        lv = str(cell.value).strip().replace(" ", "").replace("\n", "") if cell.value is not None else ""
        if lv == label.replace(" ", "").replace("\n", ""):
            return sheet[f"{value_col}{cell.row}"].value
    return None


def find_input_value_address(sheet, label: str, label_col: str = "B", value_col: str = "C", max_rows: int = 200):
    """입력 시트에서 레이블의 주소 찾기"""
    rng_labels = sheet[f"{label_col}1:{label_col}{max_rows}"]
    for cell in rng_labels:
        lv = str(cell.value).strip().replace(" ", "").replace("\n", "") if cell.value is not None else ""
        if lv == label.replace(" ", "").replace("\n", ""):
            return f"{value_col}{cell.row}"
    return None


def get_year_values(sheet, label, year_cols=("C", "D"), label_col="B"):
    """1차년도, 2차년도 값 가져오기"""
    norm_label = normalize(label)
    
    for cell in sheet.range(f"{label_col}18:{label_col}19"):
        if normalize(cell.value) == norm_label:
            row = cell.row
            year1 = sheet[f"{year_cols[0]}{row}"].value
            year2 = sheet[f"{year_cols[1]}{row}"].value
            return year1, year2
    
    raise ValueError(f"'{label}' 행을 찾지 못했습니다.")


def get_label_row(sheet, label, label_col="A", max_rows=200):
    """레이블의 행 번호 찾기"""
    norm_label = normalize(label)
    for cell in sheet.range(f"{label_col}1:{label_col}{max_rows}"):
        if normalize(cell.value) == norm_label:
            return cell.row
    raise ValueError(f"'{label}' 행을 찾지 못했습니다.")


def clear_row_range(sheet, label, label_col="A", first_value_col="B", max_cols=300, max_rows=200):
    """특정 레이블 행의 값 범위 지우기"""
    row = get_label_row(sheet, label, label_col, max_rows)
    start_col = sheet.range(f"{first_value_col}{row}").column
    # 기존 값 완전히 지우기
    sheet.range((row, start_col)).resize(1, max_cols).value = None
    sheet.range((row, start_col)).resize(1, max_cols).color = None


def set_row_values(sheet, label, values, label_col="A", first_value_col="B", max_rows=200):
    """특정 레이블 행에 값 설정"""
    row = get_label_row(sheet, label, label_col, max_rows)
    start_col = sheet.range(f"{first_value_col}{row}").column
    sheet.range((row, start_col)).resize(1, len(values)).value = [values]


def find_age_row(issue_age):
    """데이터 시트에서 나이에 해당하는 행 찾기"""
    values = da.range("A3").expand('down').value
    if isinstance(values, list) and values and isinstance(values[0], list):
        values = [row[0] for row in values]

    cleaned = []
    for val in values:
        if val is None:
            cleaned.append(None)
        else:
            match = re.search(r'\d+', str(val))
            cleaned.append(int(match.group()) if match else None)

    try:
        idx = next(i for i, v in enumerate(cleaned) if v == issue_age)
    except StopIteration:
        raise ValueError(f"데이터 시트에서 {issue_age}세 행을 찾지 못했습니다.")

    return idx + 3


# ==================== 계산 함수 ====================
def get_mortality_list(issue_age, policy_term, gender):
    """보험기간 동안의 사망률 목록 가져오기"""
    row = find_age_row(issue_age)
    col_letter = 'C' if gender == '남' else 'B'
    
    start = da.range(f"{col_letter}{row}")
    vals = start.resize(policy_term, 1).value
    
    if policy_term == 1:
        mortality_list = [vals]
    elif isinstance(vals, list) and vals and isinstance(vals[0], list):
        mortality_list = [v[0] for v in vals]
    else:
        mortality_list = vals
    
    return mortality_list


def calculate_death_benefit_pv(mortality_list, sum_assured, discount_rate):
    """사망보험금 현재가치 계산"""
    survival_prob = 1.0
    death_components = []
    
    for year, q in enumerate(mortality_list, start=0):
        power = year + 1
        term_numerator = survival_prob * q
        term_denominator = (1 + discount_rate) ** power
        term_value = term_numerator / term_denominator
        
        survival_prob *= (1 - q)
        death_components.append(term_value * sum_assured)
    
    return death_components


def pv_death_benefit_from(start_idx: int, mortality_list: List, sum_assured: float, discount_rate: float) -> float:
    """특정 시점부터의 사망보험금 현재가치"""
    future_mort = mortality_list[start_idx:]
    if not future_mort:
        return 0.0

    survival_prob = 1.0
    total = 0.0
    for offset, q in enumerate(future_mort, start=1):
        term_numerator = survival_prob * q
        term_denominator = (1 + discount_rate) ** offset
        term_value = term_numerator / term_denominator
        
        survival_prob *= (1 - q)
        total += term_value * sum_assured
    return total


def pv_level_amount_from(amount: float, start_idx: int, policy_term: int, discount_rate: float, start_power: int = 1) -> float:
    """특정 시점부터의 정액 현금흐름 현재가치"""
    periods = policy_term - start_idx
    if periods <= 0:
        return 0.0

    total = 0.0
    for offset in range(periods):
        exp = start_power + offset
        total += amount / (1 + discount_rate) ** exp
    return total


def pv_premium_with_survival(amount: float, start_idx: int, policy_term: int, mortality_list: List, discount_rate: float) -> float:
    """특정 시점부터의 보험료 현재가치 (생존율 누적 반영)"""
    periods = policy_term - start_idx
    if periods <= 0:
        return 0.0

    total = 0.0
    survival_prob = 1.0  # 초기 생존율
    
    for offset in range(periods):
        power = offset  # start_idx부터 시작하므로 0, 1, 2...
        discounted_premium = amount * survival_prob / (1 + discount_rate) ** power
        total += discounted_premium
        
        # 다음 연도를 위해 생존율 누적
        actual_year_idx = start_idx + offset
        if actual_year_idx < len(mortality_list):
            survival_prob *= (1 - mortality_list[actual_year_idx])
    
    return total


def pv_expense_with_survival(amount: float, start_idx: int, policy_term: int, mortality_list: List, discount_rate: float) -> float:
    """특정 시점부터의 사업비 현재가치 (기시 처리, 누적 생존율 반영)"""
    periods = policy_term - start_idx
    if periods <= 0:
        return 0.0

    total = 0.0
    survival_prob = 1.0  # 초기 생존율
    
    for offset in range(periods):
        power = offset  # 기시 처리이므로 0, 1, 2...
        discounted = amount * survival_prob / (1 + discount_rate) ** power
        total += discounted
        
        # 다음 연도를 위해 생존율 누적
        actual_idx = start_idx + offset
        if actual_idx < len(mortality_list):
            survival_prob *= (1 - mortality_list[actual_idx])
    
    return total


def pv_surrender_with_survival(amount: float, start_idx: int, policy_term: int, mortality_list: List, discount_rate: float) -> float:
    """특정 시점부터의 해약환급금 현재가치 (기말 처리, 누적 생존율 반영)"""
    periods = policy_term - start_idx
    if periods <= 0:
        return 0.0

    total = 0.0
    survival_prob = 1.0  # 초기 생존율
    
    for offset in range(periods):
        power = offset + 1  # 기말 처리이므로 1, 2, 3...
        actual_year_idx = start_idx + offset
        
        # 기말이므로 해당 년도 말까지의 생존율 먼저 반영 (누적)
        if actual_year_idx < len(mortality_list):
            survival_prob *= (1 - mortality_list[actual_year_idx])
        
        discounted_amount = amount * survival_prob / (1 + discount_rate) ** power
        total += discounted_amount
    
    return total




def calculate_bel_by_time(policy_term, mortality_list, annual_premium, direct_expense, 
                          surrender_value, sum_assured, discount_rate):
    """시점별 최선추정부채(BEL) 계산"""
    bel_by_time = {}
    
    for t in range(policy_term + 1):
        # 보험료는 생존율 누적 반영 (기시 처리)
        premium_pv = pv_premium_with_survival(annual_premium, t, policy_term, mortality_list, discount_rate)
        # 사업비는 생존율 누적 반영 (기시 처리)
        expense_pv = pv_expense_with_survival(direct_expense, t, policy_term, mortality_list, discount_rate)
        # 해약환급금은 생존율 누적 반영 (기말 처리)
        surrender_pv = pv_surrender_with_survival(surrender_value, t, policy_term, mortality_list, discount_rate)
        death_pv = pv_death_benefit_from(t, mortality_list, sum_assured, discount_rate)
        
        bel_t = (death_pv + expense_pv + surrender_pv) - premium_pv  # 최선추정부채(BEL) = 장래지출 - 장래수입
        bel_by_time[t] = bel_t
    
    return bel_by_time

def calculate_ra_simple(bel_t):
    """위험조정(RA) 계산"""
    confidence_level = 0.262
    risk_factors = {
        "사망 위험계수": 0.0327,
        "해약 위험계수": 0.0916,
        "사업비 위험계수": 0.0262,
    }

    # 상관행렬
    A = np.array([[1, 0, 0], [0, 1, 0], [0.25, 0.5, 1]])
    AT = A.T
    S = A @ AT
    D = np.diag(np.sqrt(np.diag(S)))
    D_inv = np.linalg.inv(D)
    R = D_inv @ S @ D_inv

    death_ra = risk_factors["사망 위험계수"] * confidence_level * bel_t
    surrender_ra = risk_factors["해약 위험계수"] * confidence_level * bel_t
    expense_ra = risk_factors["사업비 위험계수"] * confidence_level * bel_t
    
    # 총위험조정 계산
    cov_01 = R[0,1] * death_ra * surrender_ra
    cov_02 = R[0,2] * death_ra * expense_ra
    cov_12 = R[1,2] * surrender_ra * expense_ra
    
    ra_total = math.sqrt(death_ra**2 + surrender_ra**2 + expense_ra**2 + 2 * (cov_01 + cov_02 + cov_12))
    
    return ra_total


def calculate_ra_by_time(bel_by_time, policy_term):
    """시점별 실제 RA 계산 (BEL에 비례)"""
    ra_dict = {}
    
    for t in range(policy_term + 1):
        if t in bel_by_time:
            ra_dict[t] = calculate_ra_simple(bel_by_time[t])
        else:
            ra_dict[t] = 0.0
    
    return ra_dict


def calculate_ra_exp_and_diff(ra_dict, policy_term):
    """예상 RA 및 차이값 계산"""
    ra_exp_dict = {0: None}
    ra_diff_dict = {0: None}
    
    for t in range(1, policy_term + 1):
        # 예상 RA: ra_exp_t = ra_(t-1) - ra_(t-1) / (보험기간 - (t-1))
        ra_prev = ra_dict.get(t-1, 0.0)
        remaining = policy_term - (t-1)
        ra_exp_dict[t] = ra_prev * (remaining - 1) / remaining if remaining > 0 else 0.0
        
        # 차이값: ra_diff_t = ra_exp_t - ra_t
        ra_diff_dict[t] = ra_exp_dict[t] - ra_dict.get(t, 0.0)
    
    return ra_exp_dict, ra_diff_dict



def calculate_csm_series(csm_0, csm_1, csm_2, policy_term, discount_rate, ra_diff_dict):
    """CSM 시계열 계산 (RA 경험조정 반영)"""
    policy_term = int(policy_term)
    csm_prev = csm_2
    csm_series = {}
    
    for year in range(3, policy_term + 1):
        interest = csm_prev * discount_rate
        
        # 미래현금서비스 변화 = ra_diff_{year}
        future_service_change = ra_diff_dict.get(year, 0.0)
        
        numerator = csm_prev + interest + future_service_change
        denom = policy_term - (year - 1)
        
        if numerator > 0 and denom > 0:
            release = numerator / denom
            csm_curr = numerator - release
        else:
            release = 0.0
            csm_curr = 0.0
        
        csm_series[year] = csm_curr
        csm_prev = csm_curr
    
    return csm_series


# ==================== 메인 함수 ====================
def connect_excel():
    """엑셀 파일 연결"""
    global wb, inp, da, cal, out
    
    try:
        # VBA에서 호출될 때
        wb = xw.Book.caller()
        print("✓ VBA에서 호출됨")
    except:
        # 직접 실행할 때
        try:
            if xw.books:
                wb = xw.books.active
                print("✓ 활성화된 엑셀 파일 사용")
            else:
                # 상대 경로로 엑셀 파일 찾기 (스크립트와 같은 폴더)
                script_dir = os.path.dirname(os.path.abspath(__file__))
                path = os.path.join(script_dir, "IFRS17_CSM_portfolio.xlsm")
                wb = xw.Book(path)
                print("✓ 파일에서 엑셀 열기")
        except Exception as e:
            print(f"❌ 엑셀 연결 실패: {e}")
            raise
    
    inp = wb.sheets['입력']
    da = wb.sheets['데이터']
    cal = wb.sheets['계산']
    out = wb.sheets['결과출력']
    
    print("✓ 엑셀 시트 연결 성공")


def setup_output_sheet():
    """결과 출력 시트 초기 설정"""
    # 보험기간 주소 찾기
    addr = find_input_value_address(inp, "보험기간(년)")
    if addr is None:
        raise ValueError("입력 시트에서 '보험기간' 주소를 찾지 못했습니다.")
    
    # B1 기준, 오른쪽으로 보험기간 수만큼 수식 채우기
    max_cols = 400
    formula = f'=IF(COLUMN()-1<=\'{inp.name}\'!{addr}, COLUMN()-1 & "차년도", "")'
    rg = out.range("B1").resize(1, max_cols)
    rg.formula = [[formula] * max_cols]
    
    # A1부터 같은 열에 라벨 기입
    labels = [
       "PV(유입)_보험료",
        "PV(유출)_보험금",
        "PV(유출)_해약환급금",
        "PV(유출)_직접사업비",
        "최선추정부채(BEL)",
        "위험조정(RA)",
        "이행현금흐름(FCF)",
        "보험계약마진(CSM)",
        "보험계약부채",
    ]
    out.range("A1").value = [[lbl] for lbl in labels]


def get_input_parameters():
    """입력 파라미터 가져오기"""
    params = {}
    for key, label in LABELS.items():
        params[key] = get_input_value(inp, label)
    
    # 편의용 변수 추출
    issue_age = int(params["issue_age"])
    policy_term = int(params["policy_term"])
    annual_premium = params["annual_premium"]
    sum_assured = params["sum_assured"]
    discount_rate = params["discount_rate"]
    expense_ratio = params["expense_ratio"]
    direct_expense = params["direct_expense"]
    surrender_value = params["surrender_value"]
    
    # 성별
    gender = str(inp['C3'].value or '').strip()
    
    print("\n=== 입력 파라미터 ===")
    for key, value in params.items():
        print(f"{LABELS[key]}: {value}")
    print(f"성별: {gender}")
    
    return {
        'issue_age': issue_age,
        'policy_term': policy_term,
        'annual_premium': annual_premium,
        'sum_assured': sum_assured,
        'discount_rate': discount_rate,
        'expense_ratio': expense_ratio,
        'direct_expense': direct_expense,
        'surrender_value': surrender_value,
        'gender': gender,
    }


def calculate_initial_values(params):
    """초기값 계산 (BEL, RA, FCF, CSM)"""
    print("\n=== 초기값 계산 ===")
    
    # 사망률 목록
    mortality_list = get_mortality_list(params['issue_age'], params['policy_term'], params['gender'])
    print(f"사망률 목록: {mortality_list}")
    
    # 사망보험금 연도별 현재가치 (엑셀 출력용 death_components 생성)
    death_components = calculate_death_benefit_pv(
        mortality_list, params['sum_assured'], params['discount_rate']
    )
    
    # 시점별 BEL 계산 (주피터 노트북 방식)
    bel_by_time = {}
    for t in range(params['policy_term'] + 1):
        premium_pv = pv_premium_with_survival(params['annual_premium'], t, params['policy_term'], mortality_list, params['discount_rate'])
        expense_pv = pv_expense_with_survival(params['direct_expense'], t, params['policy_term'], mortality_list, params['discount_rate'])
        surrender_pv = pv_surrender_with_survival(params['surrender_value'], t, params['policy_term'], mortality_list, params['discount_rate'])
        death_pv = pv_death_benefit_from(t, mortality_list, params['sum_assured'], params['discount_rate'])
        
        bel_t = (death_pv + expense_pv + surrender_pv) - premium_pv  # BEL = 장래지출 - 장래수입
        bel_by_time[t] = bel_t
    
    bel_0 = bel_by_time[0]
    print(f"최선추정부채(BEL(0)): {bel_0:.6f}")
    
    # RA (0시점)
    ra_0 = calculate_ra_simple(bel_0)
    print(f"총위험조정(RA(0)): {ra_0:.6f}")
    
    # FCF
    fcf = bel_0 + ra_0
    print(f"이행현금흐름(FCF): {fcf:.6f}")
    
    # CSM
    csm_0 = max(-fcf, 0)
    print(f"보험계약마진(CSM(0)): {csm_0:.6f}")
    
    # 보험계약부채
    liability = fcf + csm_0
    print(f"보험계약부채: {liability:.6f}")
    
    return {
        'mortality_list': mortality_list,
        'death_components': death_components,
        'bel_by_time': bel_by_time,
        'bel_0': bel_0,
        'ra_0': ra_0,
        'fcf': fcf,
        'csm_0': csm_0,
        'liability': liability,
    }


def calculate_csm_year1_year2(params, initial_values, ra_diff_dict):
    """CSM 1차년도, 2차년도 계산 (RA 경험조정 반영)"""
    print("\n=== CSM 1-2차년도 계산 ===")
    
    # 실제 보험료, 해약환급금
    premium_actual_year1, premium_actual_year2 = get_year_values(inp, LABELS_TWO_YEARS["premium_actual"])
    surrender_actual_year1, surrender_actual_year2 = get_year_values(inp, LABELS_TWO_YEARS["surrender_actual"])
    
    csm_0 = initial_values['csm_0']
    policy_term = params['policy_term']
    discount_rate = params['discount_rate']
    annual_premium = params['annual_premium']
    surrender_value = params['surrender_value']
    
    # 1차년도
    csm_interest_1 = (csm_0 + premium_actual_year1 - annual_premium + surrender_value - surrender_actual_year1) * discount_rate
    csm_experience_1 = (premium_actual_year1 - annual_premium) + (surrender_value - surrender_actual_year1)
    csm_future_service_1 = ra_diff_dict.get(1, 0.0)  # RA 경험조정 반영
    
    if (csm_0 + csm_interest_1 + csm_experience_1 + csm_future_service_1) > 0:
        csm_release_1 = (csm_0 + csm_interest_1 + csm_experience_1 + csm_future_service_1) / policy_term
        csm_1 = csm_0 + csm_interest_1 + csm_experience_1 + csm_future_service_1 - csm_release_1
    else:
        csm_release_1 = 0
        csm_1 = 0
    
    print(f"CSM(1): {csm_1:.6f}")
    
    # 2차년도
    csm_interest_2 = (csm_1 + premium_actual_year2 - annual_premium + surrender_value - surrender_actual_year2) * discount_rate
    csm_experience_2 = (premium_actual_year2 - annual_premium) + (surrender_value - surrender_actual_year2)
    csm_future_service_2 = ra_diff_dict.get(2, 0.0)  # RA 경험조정 반영
    
    if (csm_1 + csm_interest_2 + csm_experience_2 + csm_future_service_2) > 0:
        csm_release_2 = (csm_1 + csm_interest_2 + csm_experience_2 + csm_future_service_2) / (policy_term - 1)
        csm_2 = csm_1 + csm_interest_2 + csm_experience_2 + csm_future_service_2 - csm_release_2
    else:
        csm_release_2 = 0
        csm_2 = 0
    
    print(f"CSM(2): {csm_2:.6f}")
    
    return csm_1, csm_2


def write_results_to_sheet(params, initial_values, csm_1, csm_2, csm_series, ra_dict):
    """결과를 엑셀 시트에 작성"""
    print("\n=== 결과 작성 중 ===")
    
    policy_term = int(params['policy_term'])
    
    # 보험료
    clear_row_range(out, "PV(유입)_보험료", first_value_col="B")
    premium_values = [params['annual_premium']] * policy_term
    set_row_values(out, "PV(유입)_보험료", premium_values, first_value_col="B")
    
    # 위험조정 (RA) - 시점별 RA 값 사용
    clear_row_range(out, "위험조정(RA)", first_value_col="B")
    ra_values = [ra_dict.get(t, 0.0) for t in range(policy_term)]  # 시점 0~(policy_term-1)
    set_row_values(out, "위험조정(RA)", ra_values, first_value_col="B")
    
    # 해약환급금
    clear_row_range(out, "PV(유출)_해약환급금", first_value_col="C")
    surrender_values = [params['surrender_value']] * max(policy_term, 0)
    if surrender_values:
        set_row_values(out, "PV(유출)_해약환급금", surrender_values, first_value_col="C")
    
    # 직접사업비
    clear_row_range(out, "PV(유출)_직접사업비", first_value_col="B")
    expense_values = [params['direct_expense']] * max(policy_term, 0)
    if expense_values:
        set_row_values(out, "PV(유출)_직접사업비", expense_values, first_value_col="B")
    
    # 보험금 (사망보험금)
    clear_row_range(out, "PV(유출)_보험금", first_value_col="C")
    if initial_values['death_components']:
        set_row_values(out, "PV(유출)_보험금", initial_values['death_components'], first_value_col="C")
    
    # FCF (초기시점만)
    clear_row_range(out, "이행현금흐름(FCF)", first_value_col="B", max_cols=1)
    set_row_values(out, "이행현금흐름(FCF)", [initial_values['fcf']], first_value_col="B")
    
    # BEL 시점별 (이미 계산된 값 사용)
    bel_values = [initial_values['bel_by_time'].get(t, 0.0) for t in range(policy_term + 1)]
    clear_row_range(out, "최선추정부채(BEL)", first_value_col="B")
    set_row_values(out, "최선추정부채(BEL)", bel_values, first_value_col="B")
    
    # CSM 시계열
    csm_series_full = {0: initial_values['csm_0'], 1: csm_1, 2: csm_2}
    csm_series_full.update(csm_series)  # 3차년도 이후 값 포함
    csm_values = [csm_series_full.get(t, 0.0) for t in range(policy_term + 1)]
    clear_row_range(out, "보험계약마진(CSM)", first_value_col="B")
    set_row_values(out, "보험계약마진(CSM)", csm_values, first_value_col="B")
    
    # 보험계약부채 (BEL + RA + CSM) - 시점별 RA 사용
    max_length = len(bel_values)
    liability_values = []
    for i in range(max_length):
        bel = bel_values[i] if i < len(bel_values) else 0.0
        ra = ra_dict.get(i, 0.0)  # 시점별 RA 사용
        csm = csm_values[i] if i < len(csm_values) else 0.0
        liability_values.append(bel + ra + csm)
    
    clear_row_range(out, "보험계약부채", first_value_col="B")
    set_row_values(out, "보험계약부채", liability_values, first_value_col="B")
    
    print("✓ 결과 작성 완료")


def run_calculation(sync_layout_flag: bool = True):
    """계산 실행 함수 (계산 버튼에서 호출)"""
    print("=" * 60)
    print("IFRS17 CSM 포트폴리오 계산 시작")
    print("=" * 60)
    
    # 1. 엑셀 연결
    connect_excel()
    
    # 2. 레이아웃 자동 설정 (보험기간 변경 시 자동 반영)
    if sync_layout_flag and sync_layout is not None:
        try:
            sync_layout()
            print("✓ 레이아웃 자동 설정 완료")
        except Exception as e:
            print(f"⚠ 레이아웃 설정 실패: {e}")
    
    # 3. 입력 파라미터 가져오기
    params = get_input_parameters()
    
    # 4. 초기값 계산 (시점별 BEL 포함)
    initial_values = calculate_initial_values(params)
    bel_by_time = initial_values['bel_by_time']
    
    # 5. 시점별 RA 계산
    ra_dict = calculate_ra_by_time(bel_by_time, params['policy_term'])
    print(f"시점별 RA 계산 완료")
    
    # 6. 예상 RA 및 차이값 계산
    ra_exp_dict, ra_diff_dict = calculate_ra_exp_and_diff(ra_dict, params['policy_term'])
    print(f"예상 RA 및 차이값 계산 완료")
    
    # 7. CSM 1-2차년도 계산 (경험조정(보험료, 해약환급금), 미래현금서비스 변화(RA) 반영)
    csm_1, csm_2 = calculate_csm_year1_year2(params, initial_values, ra_diff_dict)
    
    # 8. CSM 3차년도 이후 계산 (미래현금서비스 변화(RA)만 반영)
    csm_series = calculate_csm_series(initial_values['csm_0'], csm_1, csm_2, 
                                      params['policy_term'], params['discount_rate'], ra_diff_dict)
    
    # 9. 결과 작성
    write_results_to_sheet(params, initial_values, csm_1, csm_2, csm_series, ra_dict)
    
    print("\n" + "=" * 60)
    print("✓ 포트폴리오 계산 완료!")
    print("=" * 60)


def main():
    """메인 실행 함수 (하위 호환성 유지)"""
    run_calculation()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()

