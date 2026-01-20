import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)본부 한시정원(부서명 입력).xlsx"

def inspect_hq_temp():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print("\n--- Inspecting Headers (Rows 6,7,8 -> Indices 5,6,7) ---")
    # Using J to M (Indices 9 to 12)
    # Plus Column G (Index 6) for remarks
    print(df.iloc[5:8, 9:13].to_string())

    print("\n--- Inspecting Data Range (Rows 11-117 -> Indices 10-116) ---")
    # Sample first data row
    print("Row 11 Values (Sample):", df.iloc[10, 9:13].values)
    print("Column G (Index 6) Row 11:", df.iloc[10, 6])
    
    # Check for any data
    data_count = 0
    for r_idx in range(10, 117):
        row_data = df.iloc[r_idx, 9:13]
        for v in row_data:
            try:
                if float(v) > 0:
                    data_count += 1
            except: pass
    print(f"Non-zero values found in J11:M117 range: {data_count}")

if __name__ == "__main__":
    inspect_hq_temp()
