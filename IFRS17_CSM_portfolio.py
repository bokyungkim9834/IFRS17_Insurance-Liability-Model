# %%
# 라이브러리 불러오기
# 엑셀파일 열기 및 수정.저장.계산 가능
import xlwings as xw
import pandas as pd
import numpy as np
import re
from openpyxl import load_workbook

# %%
# 엑셀 파일 연결
def main():
    wb = xlwings.Book.caller()
    inp = wb.sheets["입력"]
    da = wb.sheets["데이터"]
    cal = wb.sheets["계산"]
    out = wb.sheets["결과출력"]
if __name__ == "__main__":
    path = r"C:\Users\naya9\OneDrive\바탕 화면\IFRS17_CSM_portfolio"
    xlwings.Book(path).set_mock_caller()
    main()

    # 매크로 호출 추가 
    try:
     # 보험기간 계산 먼저 반영
        wb.macro('policyterm')()
    except Exception as e:
        print("매크로 실행 실패", e)

# %%


# %%
# xlwings 유틸 함수: get_value_by_label / refresh_once / start_polling
import time
from typing import Optional, Dict
# 항목명으로 값(3번째 열)을 찾는 유틸리티
# 입력 시트에서 레이블(B열)로 값을(C열) 찾기
def get_input_value(sheet, label: str, label_col: str = "B", value_col: str = "C", max_rows: int = 200):
    rng_labels = sheet[f"{label_col}1:{label_col}{max_rows}"]
    for cell in rng_labels:
        lv = str(cell.value).strip().replace(" ", "").replace("\n", "") if cell.value is not None else ""
        if lv == label.replace(" ", "").replace("\n", ""):
            return sheet[f"{value_col}{cell.row}"].value
    return None

LABELS = {
    "issue_age": "피보험자 나이(세)",
    "policy_term": "보험기간(년)",
    "annual_premium": "연납보험료",
    "sum_assured": "보험가입금액",
    "discount_rate": "할인율",
    "lapse_rate": "해약률",
    "expense_ratio": "사업비율",
    "direct_expense": "직접사업비",
    "surrender_value": "해약환급금",
    "mortality_rate": "사망률",
}
# 입력시트에 있는 레이블을 읽어 계산하고 결과출력 시트에 기록
# 이미 연결된 inp/out을 사용해 한 번 읽고 계산/기록
def refresh_once() -> Dict[str, Optional[float]]:
    values: Dict[str, Optional[float]] = {}
    for key, label in LABELS.items():
        values[key] = get_input_value(inp, label)

    ap = values.get("annual_premium") or 0
    dr = values.get("discount_rate") or 0
    calc_result = ap * (1 + dr)
    out["B2"].value = calc_result
    return values

# 주기적 폴링
# 1초마다 입력시트에 있는 레이블을 읽어 계산하고 결과출력 시트에 기록
def start_polling(interval_sec: float = 1.0, iterations: Optional[int] = None):
    n = 0
    try:
        while True:
            vals = refresh_once()
            print({k: vals.get(k) for k in ("annual_premium", "discount_rate", "mortality_rate")})
            time.sleep(interval_sec)
            n += 1
            if iterations is not None and n >= iterations:
                break
    except KeyboardInterrupt:
        print("폴링 중단")

print("입력시트 레이블 읽기 준비 완료")

# %%
# 보험기간 값이 있는 '입력' 시트의 주소 찾기
inp = wb.sheets["입력"]
da = wb.sheets["데이터"]
cal = wb.sheets["계산"]
out = wb.sheets["결과출력"]
try:
    addr
except NameError:
    def find_input_value_address(sheet, label: str, label_col: str = "B", value_col: str = "C", max_rows: int = 200):
        rng_labels = sheet[f"{label_col}1:{label_col}{max_rows}"]
        for cell in rng_labels:
            lv = str(cell.value).strip().replace(" ", "").replace("\n", "") if cell.value is not None else ""
            if lv == label.replace(" ", "").replace("\n", ""):
                return f"{value_col}{cell.row}"
        return None
addr = find_input_value_address(inp, "보험기간(년)")

if addr is None:
    raise ValueError("입력 시트에서 '보험기간' 주소를 찾지 못했습니다.")

# B1 기준, 오른쪽으로 보험기간 수만큼 수식 채우기
max_cols = 400
formula = f'=IF(COLUMN()-1<=\'{inp.name}\'!{addr}, COLUMN()-1 & "차년도", "")'
rg = out.range("B1").resize(1, max_cols)
# 같은 수식을 전체에 채우면 COLUMN()이 각 셀 기준으로 평가되어 1,2,3...차년도 형태로 동작
rg.formula = [[formula] * max_cols]

# A1부터 같은 열에 요청 라벨 기입
labels = [
    "구분",
    "PV(유입)_보험료",
    "PV(유출)_보험금",
    "최선추정부채(BEL)",
    "위험조정(RA)",
    "이행현금흐름(FCF)",
    "보험계약마진(CSM)",
    "보험계약부채",
]
out.range("A1").value = [[lbl] for lbl in labels]


# %%
# 계산작업을 위해 입력 시트 레이블 값 일괄 가져오기 및 표시
# 입력 시트 레이블 값 일괄 가져오기 및 표시
label_to_value = {}
for key, label in LABELS.items():
    label_to_value[label] = get_input_value(inp, label)

# 편의용 변수 바인딩
# 최적가정에 따른 예상 현금흐름
issue_age = label_to_value.get("피보험자 나이(세)")
policy_term = label_to_value.get("보험기간(년)")
annual_premium = label_to_value.get("연납보험료")
sum_assured = label_to_value.get("보험가입금액")
discount_rate = label_to_value.get("할인율")
lapse_rate = label_to_value.get("해약률")
expense_ratio = label_to_value.get("사업비율")
direct_expense = label_to_value.get("직접사업비")
surrender_value = label_to_value.get("해약환급금")
mortality_rate = label_to_value.get("사망률")

for key, value in label_to_value.items():
    print(f"{key}: {value}")


# %%
# 사망률 데이터시트 
# da = book.sheets['데이터']
data_frame = da.range("A2").expand().value
df = pd.DataFrame(data_frame[1:], columns = data_frame[0])
# 필요한 정보만 추출하기 위해 앞에 index 제거
print(df.to_string(index = False))

# %%
book = xw.Book(r"C:\Users\naya9\OneDrive\바탕 화면\IFRS17_CSM-portfolio\포트폴리오(csm).xlsm")
inp = book.sheets["입력"]
da = book.sheets["데이터"]
cal = book.sheets["계산"]
out = book.sheets["결과출력"]

# 사망보험금 계산 시 사망률 참조 위해 데이터 시트에서 사망률 찾기
# 입력시트의 나이과 깉은 행을 데이터시트에서 찾기
def find_age_row(issue_age):

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

issue_age = inp['C4'].value
row = find_age_row(issue_age)
print(f"피보험자 나이(세): {issue_age}")

# 성별참조 열 찾기
gender = str(inp['C3'].value or '').strip()
col_letter = 'C' if gender == '남' else 'B' # 남자는 C열, 여자는 B열 참조
mortality = da[f"{col_letter}{row}"].value
print(f"{gender} {issue_age}세 사망률: {mortality}") # 행과 열을 종합하여 사망률 출력

# 보험기간을 고려하여 사망률 참조
policy_term = int(float(policy_term or inp['C5'].value))
start = da.range(f"{col_letter}{row}")
vals = start.resize(policy_term, 1).value

if policy_term == 1:
    mortality_list = [vals]
elif isinstance(vals, list) and vals and isinstance(vals[0], list):
    mortality_list = [v[0] for v in vals]
else:
    mortality_list = vals

print(f"보험기간 {policy_term}년 , 보험기간에 따른 사망률 목록: {mortality_list}")

# 사망률을 고려한 사망보험금 계산
survival_prob = 1.0  # 생존확률
total_death_benefit = 0.0

for year, q in enumerate(mortality_list, start=0):  # 사망률 목록 순회
    power = year + 1  # 할인 지수: 1, 2, 3, ...
    prev_terms = mortality_list[:year]  # q0 ~ q(year-1)
    numerator_factors = [f"(1 - {prev_q})" for prev_q in prev_terms]
    numerator_expr = "*".join(numerator_factors + [str(q)]) if numerator_factors else str(q)

    term_numerator = survival_prob * q                   # (1 - q0)*(1 - q1)*...*qt * qt 표현
    term_denominator = (1+discount_rate) ** power  # 할인율 공식
    term_value = term_numerator / term_denominator  # 보험금 1원에 대한 사망률을 고려한 사망보험금의 현재가치

    print(f"{power}년도 말 현재가치: {numerator_expr} / {1+discount_rate}^{power} ≈ {term_value:.6f}")

    survival_prob *= (1 - q)
    total_death_benefit += term_value*sum_assured  # 사망률을 고려한 사망보험금의 현재가치

print(f"사망률을 고려한 총 사망보험금 현재가치 ≈ {total_death_benefit:.6f}")

# %%
# 최선추정부채(BEL) 시점별 산출 (주피터 노트북 방식)
# 보험료, 사업비, 해약환급금, 사망보험금 시점별 현재가치 계산

policy_term = int(policy_term)  # 실수면 정수로 맞추기

def pv_death_benefit_from(start_idx: int) -> float: # 사망보험금 시점별 현재가치
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

def pv_premium_from(amount: float, start_idx: int) -> float: # 보험료 시점별 현재가치 (기시 처리, 생존율 누적 반영)
    periods = policy_term - start_idx
    if periods <= 0:
        return 0.0

    total = 0.0
    survival_prob = 1.0
    
    for offset in range(periods):
        power = offset  # 기시 처리이므로 0, 1, 2...
        discounted = amount * survival_prob / (1 + discount_rate) ** power
        total += discounted
        
        # 다음 연도를 위해 생존율 누적
        actual_idx = start_idx + offset
        if actual_idx < len(mortality_list):
            survival_prob *= (1 - mortality_list[actual_idx])
    
    return total

def pv_expense_from(amount: float, start_idx: int) -> float: # 사업비 시점별 현재가치 (기시 처리, 생존율 누적 반영)
    periods = policy_term - start_idx
    if periods <= 0:
        return 0.0

    total = 0.0
    survival_prob = 1.0
    
    for offset in range(periods):
        power = offset  # 기시 처리이므로 0, 1, 2...
        discounted = amount * survival_prob / (1 + discount_rate) ** power
        total += discounted
        
        # 다음 연도를 위해 생존율 누적
        actual_idx = start_idx + offset
        if actual_idx < len(mortality_list):
            survival_prob *= (1 - mortality_list[actual_idx])
    
    return total

def pv_surrender_from(amount: float, start_idx: int) -> float: # 해약환급금 시점별 현재가치 (기말 처리, 생존율 누적 반영)
    periods = policy_term - start_idx
    if periods <= 0:
        return 0.0

    total = 0.0
    survival_prob = 1.0
    
    for offset in range(periods):
        power = offset + 1  # 기말 처리이므로 1, 2, 3...
        actual_idx = start_idx + offset
        
        # 기말이므로 해당 년도 말까지의 생존율 먼저 반영
        if actual_idx < len(mortality_list):
            survival_prob *= (1 - mortality_list[actual_idx])
        
        discounted = amount * survival_prob / (1 + discount_rate) ** power
        total += discounted
    
    return total

bel_by_time = {}

for t in range(policy_term + 1):

    premium_pv = pv_premium_from(annual_premium, t)  # 보험료 시점별 현재가치 (기시)
    expense_pv = pv_expense_from(direct_expense, t)  # 사업비 시점별 현재가치 (기시)
    surrender_pv = pv_surrender_from(surrender_value, t)  # 해약환급금 시점별 현재가치 (기말)
    death_pv = pv_death_benefit_from(t)  # 사망보험금 시점별 현재가치

    bel_t = (death_pv + expense_pv + surrender_pv) - premium_pv  # 최선추정부채(BEL) = 장래지출 - 장래수입
    bel_by_time[t] = bel_t
    print(f"최선추정부채(BEL({t})) : {bel_t:.6f}")

# 초기 BEL
bel_0 = bel_by_time[0]
print(f"\n초기 최선추정부채(BEL(0)): {bel_0:.6f}")

# %%
# 위험조정(RA) 계산
# 각 위험조정 = 위험계수 * 신뢰수준(75%) * BEL
# 위험조정 : 위험요인별상관계수를 적용한 제곱합 상관집계로 분산효과 반영하여 산출) 

confidence_level = 0.262  # 미래 현금흐름 추정치 현재가치 신뢰수준 (26.2%)
risk_factors = {
    "사망 위험계수": 0.0327,     # 사망 위험계수 3.27%
    "해약 위험계수": 0.0916,   # 해약 위험계수 9.16%
    "사업비 위험계수": 0.0262,   # 사업비 위험계수 2.62%
}

risk_adjustments = {}  # 위험계수에 따른 위험조정 구하기
for key, factor in risk_factors.items():
    ra = factor * confidence_level * bel_0
    risk_adjustments[key] = ra

death_ra = risk_adjustments["사망 위험계수"]  # 사망 위험조정
surrender_ra = risk_adjustments["해약 위험계수"]  # 해약 위험조정
expense_ra = risk_adjustments["사업비 위험계수"]  # 사업비 위험조정
print("사망 위험조정:", death_ra, "해약 위험조정:", surrender_ra, "사업비 위험조정:", expense_ra)

# 상관행렬R = D^(-1) * S * D^(-1), S = A * A^T, A = 위험요인별 상관계수 행렬
import numpy as np  # 행렬연산
# 위험요인별 상관계수 행렬
A = np.array([[1,0,0],[0,1,0],[0.25,0.5,1],  # 위험요인별 상관계수 행렬
])
AT = A. T  # 전치행렬
S = A @ AT  # 비정규화 행렬
D = np.diag(np.sqrt(np.diag(S)))
D_inv = np.linalg.inv(D)
R = D_inv @ S @ D_inv  # 정규화된 상관행렬

import math  # 제곱근 계산
# 총위험조정 = 개별위험조정 * 상관행렬 * 개별위험조정^T 의 제곱근
# 상관관계를 반영해 전체 위험조정을 결합한 값
ra_total = math.sqrt(death_ra**2+ surrender_ra**2 + expense_ra**2 + 2 * (R[0,1] *death_ra * surrender_ra + R[0,2] * death_ra * expense_ra + R[1,2] * surrender_ra * expense_ra))
print(f"총위험조정(RA): {ra_total:.6f}")

# %%
# 이행현금흐름(FCF) 계산
# 이행현금흐름 = 최선추정부채(BEL) + 위험조정(RA)
fcf = bel_0 + ra_total
print(f"이행현금흐름(FCF): {fcf:.6f}")

# 보험계약마진(CSM) 계산
# 보험계약마진 = Max[-FCF, 0]
csm_0 = max(-fcf, 0)  # 0시점 보험계약마진
print(f"보험계약마진(CSM): {csm_0:.6f}")
# 보험계약부채 계산
# 보험계약부채 = FCF + CSM
insurance_contract_liability = fcf + csm_0
print(f"보험계약부채: {insurance_contract_liability:.6f}")

# %%
import re  # 공백, 줄바꿈 제거

# CSM 1차년도 말, 2차년도 말 산출
# 1차년도, 2차년도 데이터 레이블 딕셔너리
LABELS_TWO_YEARS = {
    "premium_actual": "실제 보험료",
    "surrender_actual": "실제 해약환급금",
}

def normalize(text):
    return re.sub(r"\s+", "", str(text or ""))

def get_year_values(sheet, label, year_cols=("C", "D"), label_col="B"):
    norm_label = normalize(label)

    for cell in sheet.range(f"{label_col}19:{label_col}20"):
        if normalize(cell.value) == norm_label:
            row = cell.row
            year1 = sheet[f"{year_cols[0]}{row}"].value  # 1차년도
            year2 = sheet[f"{year_cols[1]}{row}"].value  # 2차년도
            return year1, year2

    raise ValueError(f"'{label}' 행을 찾지 못했습니다.")

premium_actual_year1,   premium_actual_year2   = get_year_values(inp, LABELS_TWO_YEARS["premium_actual"]) # 실제 보험료
surrender_actual_year1,   surrender_actual_year2   = get_year_values(inp, LABELS_TWO_YEARS["surrender_actual"]) # 실제 해약환급금

# %%
# CSM 1차년도 말, 2차년도 말 산출
# 기말CSM = 기시CSM + 이자부리 +경험조정 + 미래현금서비스 변화 - 상각
# 경험조정은 보험료와 해약환급금만 반영, 미래현금서비스 변화는 0으로 가정

print(f"CSM(0): {csm_0:.6f}")
csm_interest_1 = (csm_0 + premium_actual_year1 -annual_premium) * discount_rate  # 이자부리
csm_experience_1 = (premium_actual_year1 - annual_premium) + (surrender_value - surrender_actual_year1)  # 경험조정
csm_future_service_1 = 0  # 미래현금서비스 변화는 0으로 가정

if (csm_0 + csm_interest_1 + csm_experience_1 + csm_future_service_1) > 0:  #  0보다 크면 상각
    csm_release_1 = (csm_0 + csm_interest_1 + csm_experience_1 + csm_future_service_1) / policy_term  # 잔여보장기간으로 나눔
    csm_1 = csm_0 + csm_interest_1 + csm_experience_1 + csm_future_service_1 - csm_release_1
else:
    csm_release_1 = 0  # 0보다 작으면 상각안함
    csm_1 = 0  #CSM < 0 이면 0으로 처리

print(f"CSM(1): {csm_1:.6f}")

csm_interest_2 = (csm_1 + premium_actual_year2 -annual_premium) * discount_rate
csm_experience_2 = (premium_actual_year2 - annual_premium) + (surrender_value - surrender_actual_year2)
csm_future_service_2 = 0

if (csm_1 + csm_interest_2 + csm_experience_2 + csm_future_service_2) > 0:
    csm_release_2 = (csm_1 + csm_interest_2 + csm_experience_2 + csm_future_service_2) / (policy_term -1)
    csm_2 = csm_1 + csm_interest_2 + csm_experience_2 + csm_future_service_2 - csm_release_2
else:
    csm_release_2 = 0
    csm_2 = 0

print(f"CSM(2): {csm_2:.6f}")

# %%
policy_term = int(policy_term)
csm_prev = csm_2
csm_series = {}

for year in range(3, policy_term + 1):
    interest = csm_prev * discount_rate
    numerator = csm_prev + interest
    denom = policy_term - (year - 1) 

    if numerator > 0 and denom > 0:
        release = numerator / denom
        csm_curr = numerator - release
    else:
        release = 0.0
        csm_curr = 0.0

    csm_series[year] = csm_curr
    print(f"CSM({year}): {csm_curr:.6f}")

    csm_prev = csm_curr

# %%
# 결과출력시트에 내용채우기
def get_output_value(sheet, label, label_col="A", value_col="B", max_rows=200):
    rng = sheet[f"{label_col}1:{label_col}{max_rows}"]
    normalized = label.replace(" ", "").replace("\n", "")
    for cell in rng:
        val = cell.value
        key = str(val).strip().replace(" ", "").replace("\n", "") if val is not None else ""
        if key == normalized:
            return sheet[f"{value_col}{cell.row}"].value
    return None

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


# %%
# 보험료, 위험조정은 기시처리
# 보험금, 해약환급금은 기말처리
# 이행형금흐름(FCF)는 초년도만 반영 

def normalize(text):
    return re.sub(r"\s+", "", str(text or ""))

def get_label_row(sheet, label, label_col="A", max_rows=200):
    norm_label = normalize(label)
    for cell in sheet.range(f"{label_col}1:{label_col}{max_rows}"):
        if normalize(cell.value) == norm_label:
            return cell.row
    raise ValueError(f"'{label}' 행을 찾지 못했습니다.")

def clear_row_range(sheet, label, label_col="A", first_value_col="B", max_cols=200, max_rows=200):
    row = get_label_row(sheet, label, label_col, max_rows)
    start_col = sheet.range(f"{first_value_col}{row}").column
    sheet.range((row, start_col)).resize(1, max_cols).value = [[""] * max_cols]

def set_row_values(sheet, label, values, label_col="A", first_value_col="B", max_rows=200):
    row = get_label_row(sheet, label, label_col, max_rows)
    start_col = sheet.range(f"{first_value_col}{row}").column
    sheet.range((row, start_col)).resize(1, len(values)).value = [values]

policy_term = int(policy_term)

# 보험료레이블에 기시보험료 입력
clear_row_range(out, "PV(유입)_보험료", first_value_col="B")
premium_values = [annual_premium] * policy_term
set_row_values(out, "PV(유입)_보험료", premium_values, first_value_col="B")

# 위험조정(RA)레이블에 다음연도 기시반영 (기말평가기준 측정 후 다음연도 기시반영)
clear_row_range(out, "위험조정(RA)", first_value_col="B")
ra_values = [ra_total] * policy_term
set_row_values(out, "위험조정(RA)", ra_values, first_value_col="B")

# 해약환급금레이블에 기시해약환급금 입력
clear_row_range(out, "PV(유출)_해약환급금", first_value_col="C")
surrender_values = [surrender_value] * max(policy_term, 0)
if surrender_values:
    set_row_values(out, "PV(유출)_해약환급금", surrender_values, first_value_col="C")

# 직접사업비레이블에 기시직접사업비 입력
clear_row_range(out, "PV(유출)_직접사업비", first_value_col="B")
expense_values = [direct_expense] * max(policy_term, 0)
if expense_values:
    set_row_values(out, "PV(유출)_직접사업비", expense_values, first_value_col="B")

# FCF레이블에 초기시점에 평가, 평가시점에 입력
clear_row_range(out, "이행현금흐름(FCF)", first_value_col="B", max_cols=1)
set_row_values(out, "이행현금흐름(FCF)", [fcf], first_value_col="B")


# %%
# 보험금 레이블에 연차별 사망보험금 현재가치 입력
# 위 보험금구하는 코드와 동일
death_components = []  # 연차별 사망보험금 현재가치 저장용
survival_prob = 1.0
total_death_benefit = 0.0

for year, q in enumerate(mortality_list, start=0):
    power = year + 1
    prev_terms = mortality_list[:year]
    numerator_factors = [f"(1 - {prev_q})" for prev_q in prev_terms]
    numerator_expr = "*".join(numerator_factors + [str(q)]) if prev_terms else str(q)

    term_numerator = survival_prob * q
    term_denominator = (1 + discount_rate) ** power
    term_value = term_numerator / term_denominator  

    survival_prob *= (1 - q)
    death_components.append(term_value * sum_assured)  # 연차별 실제 금액
    total_death_benefit += death_components[-1]

clear_row_range(out, "PV(유출)_보험금", first_value_col="C")  # 기존 값 비우기
if death_components:
    set_row_values(out, "PV(유출)_보험금", death_components, first_value_col="C")

# %%
policy_term = int(policy_term)

# 1) BEL 시트 반영 (이미 계산된 값 사용)
bel_values = [bel_by_time.get(t, 0.0) for t in range(policy_term + 1)]
clear_row_range(out, "최선추정부채(BEL)", first_value_col="B")
set_row_values(out, "최선추정부채(BEL)", bel_values, first_value_col="B")

# 2) CSM 시계열 구성 및 반영
csm_series_full = {0: csm_0, 1: csm_1, 2: csm_2}
csm_series_full.update(csm_series)  # 3차년도 이후 값 포함
csm_values = [csm_series_full.get(t, 0.0) for t in range(policy_term + 1)]

clear_row_range(out, "보험계약마진(CSM)", first_value_col="B")
set_row_values(out, "보험계약마진(CSM)", csm_values, first_value_col="B")



# %%
# 평가시점(B열)부터 보험기간(포함)까지 채워진 리스트라고 가정
length = min(len(bel_values), len(ra_values), len(csm_values))
liability_values = [
    bel_values[i] + ra_values[i] + csm_values[i]
    for i in range(length)
]

clear_row_range(out, "보험계약부채", first_value_col="B")
set_row_values(out, "보험계약부채", liability_values, first_value_col="B")

# 포트폴리오 완성!


