import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 기준정원(부서명 입력).xlsx"

def check_fractional():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print("\n--- Checking for Fractional Values in Rows 12-599 ---")
    
    loss_sum = 0.0
    
    # Rows 12 to 599 (Indices 11 to 598)
    for row_idx in range(11, 599):
        row_data = df.iloc[row_idx, 10:67] 
        
        # Check Dept Info
        dept_info = df.iloc[row_idx, 0:6].values
        is_empty = True
        for item in dept_info:
            if pd.notna(item) and str(item).strip() != "":
                is_empty = False
                break
        
        is_total = False
        for item in dept_info:
             if pd.notna(item) and str(item).strip() == "계":
                 is_total = True
                 break

        if not is_empty and not is_total:
            for c_idx, val in enumerate(row_data):
                if pd.notna(val):
                    try:
                        f_val = float(val)
                        if f_val % 1 != 0:
                            print(f"Row {row_idx+1}, Col Index {10+c_idx}: Value {val} (Float)")
                            loss_sum += (f_val - int(f_val))
                    except:
                        pass
                        
    print(f"Total Fractional Loss: {loss_sum}")

if __name__ == "__main__":
    check_fractional()
