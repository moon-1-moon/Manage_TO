import pandas as pd
import sys

output_file = r'c:\Workspace\Manage_TO\test_v5\(251230)기준정원(본부)_데이터.xlsx'

try:
    df = pd.read_excel(output_file)
    if len(df) > 0 and (df['기준정원'] == 1).all():
        print("VERIFICATION_SUCCESS")
    else:
        print("VERIFICATION_FAILURE")
except Exception as e:
    print(f"VERIFICATION_ERROR: {e}")
