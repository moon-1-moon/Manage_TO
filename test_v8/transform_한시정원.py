import pandas as pd
import os
import re
from pathlib import Path

# ================= 설정 부분 =================
base_dir = Path(__file__).parent
#tasks = [
#    {"file": "22년 한시정원_본부.xlsx",  "range": "L4:V85"},
#    {"file": "22년 한시정원_소속.xlsx",  "range": "L4:AE509"},
#]

#tasks = [
#    {"file": "23년 한시정원_본부.xlsx",  "range": "L4:Q85"},
#    {"file": "23년 한시정원_소속.xlsx",  "range": "L4:AA507"},
#]

#tasks = [
#    {"file": "24년 한시정원_본부.xlsx",  "range": "L4:Q86"},
#    {"file": "24년 한시정원_소속.xlsx",  "range": "L4:AA513"},
#]

tasks = [
    {"file": "25년 한시정원_본부.xlsx",  "range": "L4:O89"}, 
    {"file": "25년 한시정원_소속.xlsx",  "range": "L4:O596"}, 
]

#output_name = "22년 한시정원_데이터.xlsx"
#output_name = "23년 한시정원_데이터.xlsx"
#output_name = "24년 한시정원_데이터.xlsx"
output_name = "25년 한시정원_데이터.xlsx"
# ============================================

def col_to_idx(col_str):
    """엑셀 열 문자(A, B, AA...)를 0부터 시작하는 인덱스로 변환"""
    idx = 0
    for exp, char in enumerate(reversed(col_str.upper())):
        idx += (ord(char) - ord('A') + 1) * (26 ** exp)
    return idx - 1

def parse_range(range_str):
    """범위 문자열을 분석하여 시작/종료 행, 열 인덱스 반환"""
    start_cell, end_cell = range_str.split(':')
    s = re.match(r"([A-Z]+)([0-9]+)", start_cell.upper())
    e = re.match(r"([A-Z]+)([0-9]+)", end_cell.upper())
    s_col = col_to_idx(s.group(1));  s_row = int(s.group(2)) - 1
    e_col = col_to_idx(e.group(1));  e_row = int(e.group(2)) - 1
    return s_row, e_row, s_col, e_col

def get_output_filename(base_path, base_name):
    """파일명 중복 시 번호를 붙여 유니크한 이름 생성"""
    output_path = base_path / base_name
    if not output_path.exists():
        return output_path
    name, ext = os.path.splitext(base_name)
    counter = 2
    while True:
        new_path = base_path / f"{name}_{counter}{ext}"
        if not new_path.exists():
            return new_path
        counter += 1

def extract_date(text):
    """문자열에서 'YYYY-MM-DD' 형태의 날짜를 추출 (마지막 날짜 반환)"""
    if not isinstance(text, str):
        return ""
    matches = re.findall(r"\d{4}-\d{2}-\d{2}", text)
    return matches[-1] if matches else ""

def process_single_file(file_path, target_range):
    """개별 파일을 처리하여 데이터 행 리스트를 반환"""
    print(f"Processing: {os.path.basename(file_path)} (Range: {target_range})")

    s_row, e_row, s_col, e_col = parse_range(target_range)
    df = pd.read_excel(file_path, header=None)

    # 헤더 추출 (행 0~2)
    job_groups  = df.iloc[0, s_col:e_col+1].ffill().values
    job_grades  = df.iloc[1, s_col:e_col+1].ffill().values
    job_series  = df.iloc[2, s_col:e_col+1].values  # 3행: 날짜 포함 문자열

    # 3행에서 열별 한시_존속기한 추출
    expire_dates = [extract_date(str(v)) for v in job_series]

    rows = []

    for row_idx in range(s_row, e_row + 1):
        dept_info = df.iloc[row_idx, 0:9].values
        data_vals = df.iloc[row_idx, s_col:e_col+1].values

        for col_idx, count_val in enumerate(data_vals):
            try:
                val = float(count_val) if pd.notna(count_val) else 0.0
            except (ValueError, TypeError):
                val = 0.0

            if val > 0:
                int_part  = int(val)
                frac_part = val - int_part

                def make_row(v):
                    return list(dept_info) + [
                        job_groups[col_idx],
                        job_grades[col_idx],
                        job_series[col_idx],   # 직렬 (원본 문자열)
                        v,                      # 정원
                        1,                      # 한시 (항상 1)
                        expire_dates[col_idx],  # 한시_존속기한
                    ]

                for _ in range(int_part):
                    rows.append(make_row(1))
                if frac_part > 0.0001:
                    rows.append(make_row(frac_part))

    return rows

def main():
    all_results = []

    for task in tasks:
        file_path = base_dir / task["file"]
        if file_path.exists():
            file_rows = process_single_file(file_path, task["range"])
            all_results.extend(file_rows)
        else:
            print(f"File not found: {task['file']}")

    if not all_results:
        print("No data processed.")
        return

    columns = [
        "기관코드", "기관명", "기관구분", "소속기관차수", "중분류", "소분류", "기구유형", "보임직급", "상호이체보임직급",
        "직군", "직급", "직렬", "정원", "한시", "한시_존속기한"
    ]

    result_df = pd.DataFrame(all_results, columns=columns)
    result_df["기관코드"] = result_df["기관코드"].astype(str)
 
    final_output_path = get_output_filename(base_dir, output_name)
 
    # 기관코드 열을 엑셀에서 문자열로 강제 저장
    with pd.ExcelWriter(final_output_path, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
            for cell in row:
                cell.value = str(cell.value)
                cell.number_format = "@"  # 텍스트 서식
 
    print("-" * 40)
    print(f"Total Rows Combined : {len(result_df):,}")
    print(f"Final file saved    : {final_output_path}")

if __name__ == "__main__":
    main()