import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"
# Find latest
files = [f for f in os.listdir(base_dir) if "본부_한시정원_데이터" in f and f.endswith(".xlsx")]
files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)

if files:
    latest_file = os.path.join(base_dir, files[0])
    print(f"Checking: {latest_file}")
    df = pd.read_excel(latest_file)
    
    print(f"Rows: {len(df)}")
    print(f"Sum of '정원': {df['정원'].sum()}")
    print("\n--- Sample ---")
    print(df[["직군", "직급", "직렬", "정원", "한시", "한시_존속기한"]].to_string())
else:
    print("File not found")
