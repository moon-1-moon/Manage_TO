import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"
# Check 본부_기준정원
f1 = os.path.join(base_dir, "(251001)본부_기준정원_데이터.xlsx")
# The script appends _2 etc if exists. Let's find latest.
def get_latest(keyword):
    files = [f for f in os.listdir(base_dir) if keyword in f and f.endswith(".xlsx")]
    files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)
    if files: return os.path.join(base_dir, files[0])
    return None

f_hq_std = get_latest("본부_기준정원_데이터")
f_sov_std = get_latest("소속기관_기준정원_데이터")

print(f"Checking {f_hq_std}")
if f_hq_std:
    df = pd.read_excel(f_hq_std)
    print("Columns:", df.columns.tolist())
    if "한시" not in df.columns:
        print("PASS: '한시' column removed.")
    else:
        print("FAIL: '한시' column exists.")

print(f"Checking {f_sov_std}")
if f_sov_std:
    df = pd.read_excel(f_sov_std)
    print("Columns:", df.columns.tolist())
    if "한시" not in df.columns:
        print("PASS: '한시' column removed.")
    else:
        print("FAIL: '한시' column exists.")
