import pandas as pd
import os

# Find the latest output file
base_dir = r"d:\Workspace\Manage_TO\test_v3"
# The script output said: Saved to: d:\Workspace\Manage_TO\test_v3\(251001)___2.xlsx
# The encoding is garbled in the log, but file logic appends _2, _3 etc.
# We should look for the most recently modified file matching the pattern.

files = [f for f in os.listdir(base_dir) if "본부_기준정원_데이터" in f and f.endswith(".xlsx")]
# Sort by modification time
files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)

if not files:
    print("No output file found!")
else:
    latest_file = os.path.join(base_dir, files[0])
    print(f"Checking file: {latest_file}")
    
    df = pd.read_excel(latest_file)
    print(f"Total Rows: {len(df)}")
    print("Columns:", df.columns.tolist())
    
    print("\n--- Sample Data (First 5 Rows) ---")
    print(df.head().to_string())
    
    print("\n--- Check New Columns (Totals) ---")
    chk = df[df["총액"] == 1]
    if not chk.empty:
        print(f"Found {len(chk)} rows with '총액'. Sample:")
        print(chk[["총액", "총액_존속기한"]].head())
    else:
        print("No rows with 총액=1 found.")

    print("\n--- Check New Columns (Temp) ---")
    chk_h = df[df["한시"] == 1]
    if not chk_h.empty:
        print(f"Found {len(chk_h)} rows with '한시'. Sample:")
        print(chk_h[["한시", "한시_존속기한"]].head())
    else:
        print("No rows with 한시=1 found.")
