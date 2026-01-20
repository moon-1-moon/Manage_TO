import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 기준정원(부서명 입력).xlsx"

def analyze_sums():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    total_generated = 0
    skipped_sum = 0
    
    # Rows 12 to 599 (Indices 11 to 598)
    for row_idx in range(11, 599):
        # Calculate sum for this row (Cols K to BO -> Indices 10 to 67)
        row_data = df.iloc[row_idx, 10:67] 
        row_sum = 0
        for val in row_data:
             try:
                 v = int(val) if pd.notna(val) else 0
                 row_sum += v
             except:
                 pass
        
        # Check Dept Info (Indices 0-5)
        dept_info = df.iloc[row_idx, 0:6].values
        
        # Rule 1: Empty
        is_empty = True
        for item in dept_info:
            if pd.notna(item) and str(item).strip() != "":
                is_empty = False
                break
        
        # Rule 2: Total
        is_total_row = False
        for item in dept_info:
            if pd.notna(item) and str(item).strip() == "계":
                is_total_row = True
                break

        if not is_empty and not is_total_row:
            total_generated += row_sum
        else:
            if row_sum > 0:
                print(f"Row {row_idx+1}: Skipped (Sum={row_sum}). Empty={is_empty}, Total={is_total_row}. Dept={dept_info}")
                skipped_sum += row_sum

    print("--------------------------------------------------")
    print(f"Total Generated Sum: {total_generated}")
    print(f"Skipped Sum: {skipped_sum}")
    print(f"Total Sum (Generated + Skipped): {total_generated + skipped_sum}")

if __name__ == "__main__":
    analyze_sums()
