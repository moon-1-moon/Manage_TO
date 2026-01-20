import pandas as pd
import os
import re

# Define file paths
base_dir = r"d:\Workspace\Manage_TO\test_v3"
input_file = os.path.join(base_dir, "(251001)본부 운영정원(부서명 입력).xlsx")

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
        # Index 5 (Row 6): Job Group (직군)
        # Index 6 (Row 7): Job Grade (직급)
        # Index 7 (Row 8): Job Series (직렬)
        
        # Columns L (Index 11) to BR (Index 69) -> Slice 11:70
        
        job_groups = df.iloc[5, 11:70].ffill().values
        job_grades = df.iloc[6, 11:70].ffill().values
        job_series = df.iloc[7, 11:70].ffill().values
        
        result_rows = []
        
        # Regex patterns
        regex_total = r"총액[:\s]*(\d{4}-\d{2}-\d{2})"
        regex_temp = r"한시[:\s]*(\d{4}-\d{2}-\d{2})"
        
        # Iterate through data rows: Row 11 to 117 (Indices 10 to 116)
        # range(10, 117)
        for row_idx in range(10, 117):
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
            
            # temp_flag = 1 if "한시" in col_g_val else 0
            # temp_date = ""
            # if temp_flag:
            #     match = re.search(regex_temp, col_g_val)
            #     if match:
            #         temp_date = match.group(1)
            
            # User requested to remove "한시" columns entirely


            # Iterate through the counts columns L to BR (Indices 11 to 69)
            row_data = df.iloc[row_idx, 11:70].values
            
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
                    
                    # Create rows for integer part (value 1)
                    if int_part > 0:
                        new_row_int = [
                            dept_info[0], dept_info[1], dept_info[2], dept_info[3], dept_info[4], dept_info[5],
                            job_groups[col_idx],
                            job_grades[col_idx],
                            job_series[col_idx],
                            1, # 정원 1
                            total_flag,
                            total_date
                        ]
                        for _ in range(int_part):
                            result_rows.append(new_row_int)
                    
                    # Create row for fractional part (e.g. 0.5)
                    if frac_part > 0.0001:
                        new_row_frac = [
                            dept_info[0], dept_info[1], dept_info[2], dept_info[3], dept_info[4], dept_info[5],
                            job_groups[col_idx],
                            job_grades[col_idx],
                            job_series[col_idx],
                            frac_part, # 정원 0.5 etc
                            total_flag,
                            total_date
                        ]
                        result_rows.append(new_row_frac)

        # Create DataFrame
        columns = [
            "부서1", "부서2", "부서3", "부서4", "부서5", "부서6", 
            "직군", "직급", "직렬", "정원",
            "총액", "총액_존속기한"
        ]

        result_df = pd.DataFrame(result_rows, columns=columns)
        
        print(f"Total rows generated: {len(result_df)}")
        
        # Save to file
        output_filename = get_output_filename(base_dir, "(251001)본부_운영정원_데이터.xlsx")
        result_df.to_excel(output_filename, index=False)
        print(f"Saved to: {output_filename}")
        
    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    transform_data()
