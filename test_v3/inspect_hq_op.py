import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)본부 운영정원(부서명 입력).xlsx"

def inspect_hq_op():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print("\n--- Inspecting Headers (Rows 6,7,8 -> Indices 5,6,7) ---")
    # Using L to BR (Indices 11 to 69)
    print(df.iloc[5:8, 11:70].dropna(how='all').head().to_string())

    print("\n--- Inspecting Data Range (Rows 11-117 -> Indices 10-116) ---")
    # Sample first data row
    print("Row 11 Values (Sample):", df.iloc[10, 11:70].values)
    
    # Check for fractional values in this range
    has_fraction = False
    for r_idx in range(10, 117):
        row_vals = df.iloc[r_idx, 11:70].values
        for v in row_vals:
            if pd.notna(v) and str(v).strip() != "":
                try:
                    f = float(v)
                    if f % 1 != 0:
                        has_fraction = True
                        print(f"Fraction found at Row {r_idx+1}: {f}")
                        break
                except: pass
        if has_fraction: break
    
    if not has_fraction:
        print("No fractional values found in sample scan.")

if __name__ == "__main__":
    inspect_hq_op()
