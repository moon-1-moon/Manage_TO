import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)본부 기준정원(부서명 입력).xlsx"

def analyze_sums():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    total_sum_all = 0
    total_sum_filtered = 0
    
    # Rows 11 to 117 (Indices 10 to 116)
    for row_idx in range(10, 117):
        # Sum of cols L to BQ (Indices 11 to 68)
        row_data = df.iloc[row_idx, 11:69]
        # Clean data
        row_sum = 0
        for val in row_data:
             try:
                 v = int(val) if pd.notna(val) else 0
                 row_sum += v
             except:
                 pass
        
        total_sum_all += row_sum
        
        # Check filter
        dept_info = df.iloc[row_idx, 0:6].values
        is_empty = True
        for item in dept_info:
            if pd.notna(item) and str(item).strip() != "":
                is_empty = False
                break
        
        if not is_empty:
            total_sum_filtered += row_sum
            print(f"Row {row_idx+1}: Kept, Sum={row_sum}, Dept={dept_info}")
        else:
             print(f"Row {row_idx+1}: Skipped, Sum={row_sum} (is_empty={is_empty})")

    print("--------------------------------------------------")
    print(f"Total Sum (All Rows): {total_sum_all}")
    print(f"Total Sum (Filtered Rows): {total_sum_filtered}")

if __name__ == "__main__":
    analyze_sums()
