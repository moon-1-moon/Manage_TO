import pandas as pd
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 운영정원(부서명 입력).xlsx"

def inspect_sov_op():
    print(f"Loading file: {input_file}")
    df = pd.read_excel(input_file, header=None)
    
    print(f"Total DataFrame Shape: {df.shape}")
    
    # Check headers at Rows 6,7,8 (Indices 5,6,7)
    print("\n--- Header Sample (Indices 5,6,7) ---")
    print(df.iloc[5:8, 10:20].to_string())
    
    # Check Data Start (Row 12 -> Index 11)
    print("\n--- Data Start (Index 11) ---")
    print(df.iloc[11, 0:6].values) # Checking Dept info
    
    # Check User's Claimed End: Row 117 (Index 116)
    print("\n--- User Claimed End (Row 117 / Index 116) ---")
    print(df.iloc[116, 0:6].values)
    
    # Check Real End (Last non-empty Dept row)
    print("\n--- Finding Last Row with Dept Info ---")
    last_idx = -1
    for i in range(11, len(df)):
        dept = df.iloc[i, 0:6]
        # Check if empty
        is_empty = True
        for item in dept:
            if pd.notna(item) and str(item).strip() != "":
                is_empty = False
                break
        if not is_empty:
            last_idx = i
            
    print(f"Last non-empty Department Row Index: {last_idx} (Row {last_idx+1})")

if __name__ == "__main__":
    inspect_sov_op()
