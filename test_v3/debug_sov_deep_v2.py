import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 기준정원(부서명 입력).xlsx"

def deep_inspect_v2():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print("\n--- Inspecting Row 10 (Index 9) and Row 11 (Index 10) ---")
    # Columns K to BO (Indices 10 to 67)
    for idx in range(9, 11):
        row_data = df.iloc[idx, 10:67].values
        s = 0
        for v in row_data:
            try: s += int(v) if pd.notna(v) else 0
            except: pass
        print(f"Row {idx+1} Sum: {s}")
        print(f"Row {idx+1} Values: {row_data}")

    print("\n--- Checking for Skipped Rows with Sum == 3 or Sum <= 5 ---")
    # Rows 12 to 599 (Indices 11 to 598)
    for row_idx in range(11, 599):
        # Calculate sum
        row_data = df.iloc[row_idx, 10:67] 
        row_sum = 0
        for val in row_data:
             try:
                 v = int(val) if pd.notna(val) else 0
                 row_sum += v
             except:
                 pass
        
        dept_info = df.iloc[row_idx, 0:6].values
        is_empty = True
        for item in dept_info:
            if pd.notna(item) and str(item).strip() != "":
                is_empty = False
                break
        
        if is_empty and row_sum > 0:
            if row_sum <= 5:
                print(f"FOUND SMALL SKIPPED ROW: Row {row_idx+1} Sum={row_sum}. Dept={dept_info}")

if __name__ == "__main__":
    deep_inspect_v2()
