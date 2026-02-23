import pandas as pd
import os
import re

# ================= 설정 부분 =================
base_dir = r"d:\Workspace\Manage_TO\test_v7"
input_file = os.path.join(base_dir, "251230_기준정원_본부.xlsx")
# 작업하고 싶은 엑셀 구역 입력
target_range = "N4:BS89" 
# ============================================

def col_to_idx(col_str):
    """엑셀 열 문자(A, B, AA...)를 0부터 시작하는 인덱스로 변환"""
    exp = 0
    idx = 0
    for char in reversed(col_str.upper()):
        idx += (ord(char) - ord('A') + 1) * (26 ** exp)
        exp += 1
    return idx - 1

def parse_range(range_str):
    """'N4:BS89' 문자열을 분석하여 시작/종료 행, 열 인덱스 반환"""
    start_cell, end_cell = range_str.split(':')
    
    # 정규표현식으로 문자와 숫자 분리
    start_match = re.match(r"([A-Z]+)([0-9]+)", start_cell.upper())
    end_match = re.match(r"([A-Z]+)([0-9]+)", end_cell.upper())
    
    start_col = col_to_idx(start_match.group(1))
    start_row = int(start_match.group(2)) - 1  # 0-indexed
    
    end_col = col_to_idx(end_match.group(1))
    end_row = int(end_match.group(2)) - 1      # 0-indexed
    
    return start_row, end_row, start_col, end_col

def get_output_filename(base_path, base_name):
    output_path = os.path.join(base_path, base_name)
    if not os.path.exists(output_path): return output_path
    name, ext = os.path.splitext(base_name)
    counter = 2
    while True:
        new_name = f"{name}_{counter}{ext}"
        new_path = os.path.join(base_path, new_name)
        if not os.path.exists(new_path): return new_path
        counter += 1

def transform_data():
    print(f"Loading file: {input_file}")
    
    try:
        # 범위 해석
        s_row, e_row, s_col, e_col = parse_range(target_range)
        
        # 엑셀 로드
        df = pd.read_excel(input_file, header=None)
        
        # 헤더 정보 (기존 1, 2, 3행 유지, 열은 타겟 범위에 맞춤)
        # s_col:e_col+1 은 입력한 범위(N~BS)의 열들을 가져옴
        job_groups = df.iloc[0, s_col:e_col+1].ffill().values
        job_grades = df.iloc[1, s_col:e_col+1].ffill().values
        job_series = df.iloc[2, s_col:e_col+1].ffill().values
        
        result_rows = []
        regex_total = r"총액[:\s]*(\d{4}-\d{2}-\d{2})"
        
        # 지정한 행 범위(s_row ~ e_row) 반복
        for row_idx in range(s_row, e_row + 1):
            # 부서 정보 (A~I열)
            dept_info = df.iloc[row_idx, 0:9].values

            # [수정포인트] B열(Index 1)에서 총액 및 날짜 정보 추출
            col_b_val = str(df.iloc[row_idx, 1]) if pd.notna(df.iloc[row_idx, 1]) else ""
            
            total_flag = 1 if "총액" in col_b_val else 0
            total_date = ""
            if total_flag:
                match = re.search(regex_total, col_b_val)
                if match:
                    total_date = match.group(1)       

            # 지정한 열 범위 데이터 추출
            row_data = df.iloc[row_idx, s_col:e_col+1].values
            
            for col_idx, count_val in enumerate(row_data):
                val = 0.0
                try:
                    if pd.notna(count_val): val = float(count_val)
                except ValueError: val = 0.0
                
                if val > 0:
                    int_part = int(val)
                    frac_part = val - int_part
                    
                    # 공통 행 데이터 생성 함수화 (가독성)
                    def create_row(v):
                        return [
                            dept_info[0], dept_info[1], dept_info[2], dept_info[3], dept_info[4], 
                            dept_info[5], dept_info[6], dept_info[7], dept_info[8],
                            job_groups[col_idx], job_grades[col_idx], job_series[col_idx],
                            v, total_flag, total_date
                        ]

                    if int_part > 0:
                        for _ in range(int_part):
                            result_rows.append(create_row(1))
                    
                    if frac_part > 0.0001:
                        result_rows.append(create_row(frac_part))

        columns = [
            "기관코드", "기관명", "기관구분", "소속기관차수", "중분류", "소분류", "기구유형", "보임직급", "상호이체보임직급", 
            "직군", "직급", "직렬", "정원", "총액", "총액_존속기한"
        ]

        result_df = pd.DataFrame(result_rows, columns=columns)
        print(f"Total rows generated: {len(result_df)}")
        
        output_filename = get_output_filename(base_dir, "기준정원_본부_데이터.xlsx")
        result_df.to_excel(output_filename, index=False)
        print(f"Saved to: {output_filename}")
        
    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    transform_data()