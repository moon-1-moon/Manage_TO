import pandas as pd

output_file = r'c:\Workspace\Manage_TO\test_v5\(251230)기준정원(본부)_데이터.xlsx'

try:
    df = pd.read_excel(output_file)
    print("Columns:", df.columns.tolist())
    print("\nFirst 5 rows:")
    print(df.head().to_string())
    print(f"\nTotal rows: {len(df)}")
    
    # Check if '기준정원' is all 1s
    if (df['기준정원'] == 1).all():
        print("\nVerification Passed: '기준정원' column contains only 1s.")
    else:
        print("\nVerification Failed: '기준정원' column contains values other than 1.")
        
except Exception as e:
    print(f"Error reading output file: {e}")
