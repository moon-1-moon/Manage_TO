import pandas as pd
import re

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)본부 기준정원(부서명 입력).xlsx"

def inspect_column_g():
    df = pd.read_excel(input_file, header=None)
    # Column G is index 6
    col_g = df.iloc[10:117, 6].dropna().unique()
    
    print("Unique values in Column G:")
    for val in col_g:
        print(f"'{val}'")
        
    # Test Regex
    regex_total = r"총액[:\s]*(\d{4}-\d{2}-\d{2})"
    regex_temp = r"한시[:\s]*(\d{4}-\d{2}-\d{2})"
    
    print("\n--- Regex Test ---")
    for val in col_g:
        s_val = str(val)
        t_match = re.search(regex_total, s_val)
        h_match = re.search(regex_temp, s_val)
        
        t_flag = 1 if "총액" in s_val else 0
        h_flag = 1 if "한시" in s_val else 0
        
        t_date = t_match.group(1) if t_match else ""
        h_date = h_match.group(1) if h_match else ""
        
        if t_flag or h_flag:
            print(f"Original: {val} -> Total: {t_flag}, Date: {t_date} | Temp: {h_flag}, Date: {h_date}")

if __name__ == "__main__":
    inspect_column_g()
