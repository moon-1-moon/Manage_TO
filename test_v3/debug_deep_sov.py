import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 기준정원(부서명 입력).xlsx"

def deep_inspect():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    # 1. Check for non-numeric values in data range
    print("\n--- Non-Numeric Data Values ---")
    data_slice = df.iloc[11:599, 10:67]
    for r_idx, row in data_slice.iterrows():
        for c_idx, val in row.items():
            if pd.notna(val) and val != 0:
                try:
                    int(val)
                except:
                    print(f"Row {r_idx+1}, Col {c_idx}: Non-integer value '{val}'")

    # 2. Check for rows that are NOT empty but Skipped (likely 'Total')
    print("\n--- Rows with Data but Skipped (e.g. Total) ---")
    for row_idx in range(11, 599):
        dept_info = df.iloc[row_idx, 0:6].values
        
        is_empty = True
        for item in dept_info:
            if pd.notna(item) and str(item).strip() != "":
                is_empty = False
                break
        
        is_total_row = False
        for item in dept_info:
            if pd.notna(item) and str(item).strip() == "계":
                is_total_row = True
                break
        
        # Calculate row sum
        row_data = df.iloc[row_idx, 10:67]
        row_sum = 0
        for val in row_data:
             try: row_sum += int(val) if pd.notna(val) else 0
             except: pass
        
        if not is_empty and is_total_row:
             print(f"Row {row_idx+1}: Skipped as TOTAL. Sum={row_sum}. Dept={dept_info}")

    # 3. Check sums again with strict print
    print("\n--- Summary ---")
    generated_count = 0
    for row_idx in range(11, 599):
        dept_info = df.iloc[row_idx, 0:6].values
        
        is_empty = True
        for item in dept_info:
            if pd.notna(item) and str(item).strip() != "":
                is_empty = False
                break
                
        is_total_row = False
        for item in dept_info:
            if pd.notna(item) and str(item).strip() == "계":
                is_total_row = True
                break
                
        row_data = df.iloc[row_idx, 10:67]
        row_sum = 0
        for val in row_data:
             try: row_sum += int(val) if pd.notna(val) else 0
             except: pass
             
        if not is_empty and not is_total_row:
            generated_count += row_sum
            
    print(f"Calculated Generated Count: {generated_count}")

if __name__ == "__main__":
    deep_inspect()
