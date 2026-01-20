import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 한시정원(부서명 입력).xlsx"

def inspect_sov_temp():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print("\n--- Inspecting Headers (Rows 6,7,8 -> Indices 5,6,7) ---")
    # Using J to T (Indices 9 to 19)
    # Plus Column G (Index 6) for remarks
    print(df.iloc[5:8, 9:20].to_string())

    print("\n--- Inspecting Data Range (Rows 12-599 -> Indices 11-598) ---")
    # Sample first data row (Row 12 -> Index 11)
    print("Row 12 Values (Sample):", df.iloc[11, 9:20].values)
    print("Column G (Index 6) Row 12:", df.iloc[11, 6])
    
    # Check for any data
    data_count = 0
    quota_sum = 0.0
    for r_idx in range(11, 599):
        # Indices 9 to 19 (exclusive of 20) -> J to T (11 cols)
        row_data = df.iloc[r_idx, 9:20]
        for v in row_data:
            try:
                val = float(v)
                if val > 0:
                    data_count += 1
                    quota_sum += val
            except: pass
            
    print(f"Non-zero cells found in J12:T599 range: {data_count}")
    print(f"Sum of values: {quota_sum}")

if __name__ == "__main__":
    inspect_sov_temp()
