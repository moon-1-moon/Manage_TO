import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"
# Find the latest file (should be _3.xlsx or similar now, or at least timestamp based)
files = [f for f in os.listdir(base_dir) if "소속기관_기준정원_데이터" in f and f.endswith(".xlsx")]
files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)

if files:
    latest_file = os.path.join(base_dir, files[0])
    print(f"Checking: {latest_file}")
    df = pd.read_excel(latest_file)
    
    total_rows = len(df)
    quota_sum = df["정원"].sum()
    
    print(f"Total Rows: {total_rows}")
    print(f"Sum of '정원': {quota_sum}")
    
    # Check for 0.5s
    frac_rows = df[df["정원"] < 1]
    print(f"Fractional Rows: {len(frac_rows)}")
    print(frac_rows["정원"].value_counts())
else:
    print("No file found")
