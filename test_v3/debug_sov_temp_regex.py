import pandas as pd
import re
import os

input_file = r"d:\Workspace\Manage_TO\test_v3\(251001)소속기관 한시정원(부서명 입력).xlsx"

def debug_regex():
    df = pd.read_excel(input_file, header=None)
    # Series Row is Index 7 (Row 8)
    # Range J (9) to T (19)
    # The user says J to T. Let's look at row 8 values in that range.
    
    series_headers = df.iloc[7, 9:20].ffill().values
    
    # Regex attempt
    # We want DATE at the end: (YYYY-MM-DD)
    # We want TASK before that: (Something)
    # And the rest is CLEAN SERIES
    
    # Pattern: Look for parens with date at end.
    # Pattern: \s*\((\d{4}-\d{2}-\d{2})\)$  -> Capture date
    # Then look at what remains. If it ends with (...), that's task.
    
    regex_date = r"\s*\(\s*(\d{4}-\d{2}-\d{2})\s*\)$"
    regex_task = r"\s*\(([^)]+)\)$"
    
    print("\n--- Testing Parsing ---")
    for val in series_headers:
        if pd.isna(val): continue
        s_val = str(val).strip()
        print(f"Original: '{s_val}'")
        
        # 1. Extract Date
        date_str = ""
        clean_s = s_val
        match_date = re.search(regex_date, s_val)
        if match_date:
            date_str = match_date.group(1)
            # Remove date part
            clean_s = re.sub(regex_date, "", s_val)
            
        # 2. Extract Task (if exists in remaining)
        task_str = ""
        match_task = re.search(regex_task, clean_s)
        if match_task:
            task_str = match_task.group(1)
            # Remove task part
            clean_s = re.sub(regex_task, "", clean_s)
            
        clean_s = clean_s.strip()
        
        print(f"  -> Series: '{clean_s}' | Task: '{task_str}' | Date: '{date_str}'")

if __name__ == "__main__":
    debug_regex()
