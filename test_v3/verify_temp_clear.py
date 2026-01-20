import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"
# Find latest 본부_기준정원 output
files = [f for f in os.listdir(base_dir) if "본부_기준정원_데이터" in f and f.endswith(".xlsx")]
files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)

if files:
    latest_file = os.path.join(base_dir, files[0])
    print(f"Checking: {latest_file}")
    df = pd.read_excel(latest_file)
    
    # Check "한시" and "한시_존속기한"
    temp_counts = df["한시"].notna().sum()
    temp_date_counts = df["한시_존속기한"].notna().sum()
    
    # In regex usage, empty string might be used if regex extraction returns empty string.
    # But pandas reads empty string or NaN as NaN usually, unless strictly string.
    # Let's check value counts
    print(f"Not-NaN count in '한시': {temp_counts}")
    print(f"Not-NaN count in '한시_존속기한': {temp_date_counts}")
    
    print("\n--- Value Counts ---")
    print(df["한시"].value_counts(dropna=False))
    print(df["한시_존속기한"].value_counts(dropna=False))
else:
    print("File not found")
