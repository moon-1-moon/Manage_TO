import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"
files = [f for f in os.listdir(base_dir) if "소속기관_한시정원_데이터" in f and f.endswith(".xlsx")]
files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)

if files:
    latest_file = os.path.join(base_dir, files[0])
    print(f"Checking: {latest_file}")
    df = pd.read_excel(latest_file)
    
    print("\n--- Rows with non-empty Task (한시_업무) ---")
    task_rows = df[df["한시_업무"].notna()]
    print(f"Count: {len(task_rows)}")
    if not task_rows.empty:
        print(task_rows[["직렬", "한시_존속기한", "한시_업무"]].head(10).to_string())
    
    print("\n--- Sample of Normal Rows ---")
    print(df[["직렬", "한시_존속기한", "한시_업무"]].head(10).to_string())
else:
    print("File not found")
