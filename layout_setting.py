import xlwings as xw
from datetime import datetime


LIGHT_GREEN = (226, 239, 218)


def _to_int_term(raw_value, issue_age, limit_age: int = 100) -> int:
    # '보험기간' 셀 값을 정수 기간으로 변환
    # - 값이 "종신"이면: limit_age - issue_age 계산
    # - 그 외에는 숫자로 변환(문자 "20", "20.0" 등 허용)
    text = str(raw_value).strip() if raw_value is not None else ""
    if not text:
        raise ValueError("보험기간 값이 비어 있습니다.")
    if text.lower() == "종신":
        if issue_age is None:
            raise ValueError("종신 계산을 위해 'C4'에 연령이 필요합니다.")
        return int(limit_age - int(float(issue_age)))
    return int(float(text))

def sync_layout() -> None:
    # '입력'!C5(보험기간)을 기준으로 '결과출력' 시트 레이아웃 갱신
    # xlwings를 통해 VBA에서 호출:
    #   RunPython "import mymodule; mymodule.sync_layout()"
    wb = xw.Book.caller()
    app = wb.app

    inp = wb.sheets["입력"]
    out = wb.sheets["결과출력"]

    # 디버그 마커: 파이썬 호출 시점과 C5 값을 눈으로 확인
    out["Z1"].value = f"PY {datetime.now():%Y-%m-%d %H:%M:%S}"
    out["Z2"].value = f"C5={inp['C5'].value}"

    # 읽기 전에 엑셀 계산 강제 갱신
    try:
        app.calculate()
    except Exception:
        pass

    issue_age = inp["C4"].value
    raw_term = inp["C5"].value
    term = _to_int_term(raw_term, issue_age)
    if term <= 0:
        raise ValueError("보험기간은 1 이상이어야 합니다.")

    # A열 라벨 목록
    labels = [
        "구분(단위:백만)",
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

    num_rows = len(labels)
    num_cols = 1 + term  # A열 포함 총 열 수

    # 기존 데이터 완전히 제거 (오래된 라벨과 데이터 제거)
    _CLEAR_ROWS = 50
    _CLEAR_COLS = 300
    out["A1"].resize(_CLEAR_ROWS, _CLEAR_COLS).value = None  # 전체 영역 지우기
    out["A1"].resize(_CLEAR_ROWS, _CLEAR_COLS).color = None  # 색상도 제거
    
    # 라벨/헤더
    out["A1"].resize(num_rows, 1).value = [[label] for label in labels]
    out["A1"].resize(num_rows, 1).color = LIGHT_GREEN   

    header = ["평가시점"] + [f"{i}차년도" for i in range(1, term + 1)]
    out["B1"].resize(1, len(header)).value = [header]

    # 헤더 색상
    out["B1"].resize(1, len(header)).color = LIGHT_GREEN

    # 잔여 헤더 지우기
    _EXTRA = 200
    out["B1"].offset(0, len(header)).resize(1, _EXTRA).value = [[""] * _EXTRA]
    out["B1"].offset(0, len(header)).resize(1, _EXTRA).color = None

    # 표 범위 계산
    num_cols = 1 + len(header)  # A열 포함
    table = out["A1"].resize(num_rows, num_cols)

    xlContinuous, xlThin = 1, 2
    xlLineStyleNone = -4142
    # 표 전체에 테두리 적용 (외곽선 + 내부선)
    # 먼저 모든 테두리 적용
    for idx in (7, 8, 9, 10, 11, 12):
        try:
            b = table.api.Borders(idx)
            b.LineStyle = xlContinuous
            b.Weight = xlThin
        except Exception:
            pass
    
    # 각 셀에 개별적으로 테두리 적용 (더 확실하게)
    for row in range(1, num_rows + 1):
        for col in range(1, num_cols + 1):
            cell = out.cells(row, col)
            try:
                for border_idx in (7, 8, 9, 10):
                    b = cell.api.Borders(border_idx)
                    b.LineStyle = xlContinuous
                    b.Weight = xlThin
            except Exception:
                pass

    # 우측 잔여 영역의 테두리를 제거(이전 테두리가 남지 않도록)
    _EXTRA = 200
    residual = out["A1"].offset(0, num_cols).resize(num_rows, _EXTRA)
    # 중요: 잔여 영역의 왼쪽 테두리(7)는 표의 오른쪽 테두리와 공유되므로 지우지 않음
    for idx in (8, 9, 10, 11, 12):
        try:
            residual.api.Borders(idx).LineStyle = xlLineStyleNone
        except Exception:
            pass
    # 선택: 최종 숫자 기간을 C5에 반영(초기 입력이 '종신'이었던 경우 유용)
    inp["C5"].value = term


