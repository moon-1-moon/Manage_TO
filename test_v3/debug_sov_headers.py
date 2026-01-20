import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 기준정원(부서명 입력).xlsx"

def inspect_headers():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print("\n--- Inspecting Rows 4-10 (Indices 3-9) ---")
    # We want to find "일반직" (Job Group), "행정" (Series), "4급" (Grade)
    # Print the values in Column K (Index 10) to Column O (Index 14) for these rows
    # to identifying where the text is.
    
    for row_idx in range(3, 10):
        print(f"Row {row_idx+1} (Index {row_idx}) Sample vals (Cols 10-20):")
        print(df.iloc[row_idx, 10:20].values)

if __name__ == "__main__":
    inspect_headers()
