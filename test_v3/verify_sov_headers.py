import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"
files = [f for f in os.listdir(base_dir) if "소속기관_기준정원_데이터" in f and f.endswith(".xlsx")]
files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)

if files:
    latest_file = os.path.join(base_dir, files[0])
    print(f"Checking file: {latest_file}")
    df = pd.read_excel(latest_file)
    
    print("\n--- Rows 1-5 Sample ---")
    print(df.head().to_string())
    
    print("\n--- Unique Values in Job Columns ---")
    print("Job Group (직군):", df["직군"].unique()[:5])
    print("Job Grade (직급):", df["직급"].unique()[:5])
    print("Job Series (직렬):", df["직렬"].unique()[:5])
else:
    print("No output file found")
