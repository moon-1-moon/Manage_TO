import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 기준정원(부서명 입력).xlsx"

def inspect_file():
    print(f"Loading: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print("\n--- Rows 5-10 (Indices 4-9) Inspection ---")
    # Verify Headers at expected Rows 6, 7, 8 (Indices 5, 6, 7)
    print(df.iloc[4:10].to_string())
    
    print("\n--- Column Check ---")
    # Check Column K (Index 10) data
    # Check Column G (Index 6) remarks
    print("Row 12 (Index 11) - First Data Row Sample:")
    print(df.iloc[11].values)
    
    print("\n--- Column G Sample ---")
    print(df.iloc[11:20, 6].values)

if __name__ == "__main__":
    inspect_file()
