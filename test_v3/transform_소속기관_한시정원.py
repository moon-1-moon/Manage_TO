import pandas as pd
import os
import re

# Define file paths
base_dir = r"d:\Workspace\Manage_TO\test_v3"
input_file = os.path.join(base_dir, "(251001)소속기관 한시정원(부서명 입력).xlsx")

def get_output_filename(base_path, base_name):
    """Generates a unique filename by appending _2, _3, etc. if the file exists."""
    output_path = os.path.join(base_path, base_name)
    if not os.path.exists(output_path):
        return output_path
    
    name, ext = os.path.splitext(base_name)
    counter = 2
    while True:
        new_name = f"{name}_{counter}{ext}"
        new_path = os.path.join(base_path, new_name)
        if not os.path.exists(new_path):
            return new_path
        counter += 1

def transform_data():
    print(f"Loading file: {input_file}")
    
    try:
        # Load the entire sheet
        df = pd.read_excel(input_file, header=None)
        
        # Header Information extraction
        # Index 5 (Row 6): Job Group
        # Index 6 (Row 7): Job Grade
        # Index 7 (Row 8): Job Series
        
        # Columns J (Index 9) to T (Index 19) -> Slice 9:20
        # Wait, Range J12 to T599.
        # J is Index 9. T is 20th letter (A=0, J=9, T=19). So index 19.
        # Slice 9:20 covers indices 9 to 19 inclusive. Correct.
        
        job_groups = df.iloc[5, 9:20].ffill().values
        job_grades = df.iloc[6, 9:20].ffill().values
        raw_vocab_series = df.iloc[7, 9:20].ffill().values
        
        # Parse Series Headers
        clean_series_list = []
        series_dates = []
        series_tasks = []
        
        regex_date = r"\s*\(\s*(\d{4}-\d{2}-\d{2})\s*\)$"
        regex_task = r"\s*\(([^)]+)\)$"
        
        for val in raw_vocab_series:
            s_val = str(val).strip() if pd.notna(val) else ""
            
            # 1. Date
            d_str = ""
            clean_s = s_val
            match_date = re.search(regex_date, s_val)
            if match_date:
                d_str = match_date.group(1)
                clean_s = re.sub(regex_date, "", s_val)
            
            # 2. Task
            t_str = ""
            match_task = re.search(regex_task, clean_s)
            if match_task:
                t_str = match_task.group(1)
                clean_s = re.sub(regex_task, "", clean_s)
            
            clean_s = clean_s.strip()
            
            clean_series_list.append(clean_s)
            series_dates.append(d_str)
            series_tasks.append(t_str)
        
        result_rows = []
        
        # Regex patterns columns G
        regex_total = r"총액[:\s]*(\d{4}-\d{2}-\d{2})"
        # regex_temp = r"한시[:\s]*(\d{4}-\d{2}-\d{2})" # Not used for date anymore, using header date
        
        # Iterate through data rows: Row 12 to 599 (Indices 11 to 598)
        for row_idx in range(11, 599):
            # Check Department info in Columns A-F (Indices 0-5)
            dept_info = df.iloc[row_idx, 0:6].values
            
            # Rule 1: Skip if all empty
            is_empty = True
            for item in dept_info:
                if pd.notna(item) and str(item).strip() != "":
                    is_empty = False
                    break
            
            if is_empty:
                continue

            # Rule 2: Skip 'Total' (계) in Department columns
            is_total_row = False
            for item in dept_info:
                if pd.notna(item) and str(item).strip() == "계":
                    is_total_row = True
                    break
            
            if is_total_row:
                continue
            
            # Parse Column G (Index 6)
            col_g_val = str(df.iloc[row_idx, 6]) if pd.notna(df.iloc[row_idx, 6]) else ""
            
            total_flag = 1 if "총액" in col_g_val else 0
            total_date = ""
            if total_flag:
                match = re.search(regex_total, col_g_val)
                if match:
                    total_date = match.group(1)
            
            # Iterate through the counts columns J to T (Indices 9 to 20)
            row_data = df.iloc[row_idx, 9:20].values
            
            for col_idx, count_val in enumerate(row_data):
                if pd.isna(count_val):
                    val = 0.0
                else:
                    try:
                        val = float(count_val)
                    except ValueError:
                        val = 0.0
                
                if val > 0:
                    int_part = int(val)
                    frac_part = val - int_part
                    
                    # Force Temp = 1
                    force_temp = 1
                    
                    # Use parsed Header info
                    current_series = clean_series_list[col_idx]
                    current_temp_date = series_dates[col_idx]
                    current_temp_task = series_tasks[col_idx]
                    
                    # Create rows for integer part (value 1)
                    if int_part > 0:
                        new_row_int = [
                            dept_info[0], dept_info[1], dept_info[2], dept_info[3], dept_info[4], dept_info[5],
                            job_groups[col_idx],
                            job_grades[col_idx],
                            current_series,
                            1, # 정원 1
                            total_flag,
                            total_date,
                            force_temp,
                            current_temp_date,
                            current_temp_task
                        ]
                        for _ in range(int_part):
                            result_rows.append(new_row_int)
                    
                    # Create row for fractional part (e.g. 0.5)
                    if frac_part > 0.0001:
                        new_row_frac = [
                            dept_info[0], dept_info[1], dept_info[2], dept_info[3], dept_info[4], dept_info[5],
                            job_groups[col_idx],
                            job_grades[col_idx],
                            current_series,
                            frac_part, # 정원 0.5 etc
                            total_flag,
                            total_date,
                            force_temp,
                            current_temp_date,
                            current_temp_task
                        ]
                        result_rows.append(new_row_frac)

        # Create DataFrame
        columns = [
            "부서1", "부서2", "부서3", "부서4", "부서5", "부서6", 
            "직군", "직급", "직렬", "정원",
            "총액", "총액_존속기한", "한시", "한시_존속기한", "한시_업무"
        ]
        result_df = pd.DataFrame(result_rows, columns=columns)
        
        print(f"Total rows generated: {len(result_df)}")
        print(f"Total Quota Sum: {result_df['정원'].sum()}")
        
        # Save to file
        output_filename = get_output_filename(base_dir, "(251001)소속기관_한시정원_데이터.xlsx")
        result_df.to_excel(output_filename, index=False)
        print(f"Saved to: {output_filename}")
        
    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    transform_data()
