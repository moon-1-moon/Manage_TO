import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"

def get_latest_file(base_dir, keyword):
    files = [f for f in os.listdir(base_dir) if keyword in f and f.endswith(".xlsx")]
    # Avoid picking the result of a previous run of THIS script if any
    files = [f for f in files if "합본" not in f] 
    files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)
    if files:
        return os.path.join(base_dir, files[0])
    return None

def main():
    # 1. Identify Input Files
    op_file = get_latest_file(base_dir, "본부_운영정원_데이터")
    temp_file = get_latest_file(base_dir, "본부_한시정원_데이터")
    
    if not op_file or not temp_file:
        print(f"Error: Could not find input files.\nOp: {op_file}\nTemp: {temp_file}")
        return

    print(f"Loading Operational: {op_file}")
    df_op = pd.read_excel(op_file)
    print(f"Rows: {len(df_op)}")

    print(f"Loading Temporary: {temp_file}")
    df_temp = pd.read_excel(temp_file)
    print(f"Rows: {len(df_temp)}")
    
    # 2. Define Join Keys
    join_keys = [
        "부서1", "부서2", "부서3", "부서4", "부서5", "부서6", 
        "직군", "직급", "직렬", "정원"
    ]
    
    # 3. Perform Concatenation (Union)
    # User requested simple addition of rows.
    # Op file (652 rows) + Temp file (6 rows) = 658 rows.
    # Columns in Op: ..., 정원, 총액, 총액_존속기한
    # Columns in Temp: ..., 정원, 총액, 총액_존속기한, 한시, 한시_존속기한, 한시_업무
    
    print("Concatenating files...")
    # sort=False preserves column order as much as possible
    final_df = pd.concat([df_op, df_temp], ignore_index=True, sort=False)
    
    # 4. Verification
    print(f"Total Rows: {len(final_df)}")
    expected_rows = len(df_op) + len(df_temp)
    if len(final_df) != expected_rows:
        print(f"[WARNING] Row count mismatch! Expected {expected_rows}, got {len(final_df)}")
    
    # Check if new columns exist
    new_cols = ["한시", "한시_존속기한", "한시_업무"]
    for c in new_cols:
        if c not in final_df.columns:
            print(f"[ERROR] Column {c} missing in result.")
    
    # 5. Save
    output_filename = os.path.join(base_dir, "(251001)본부_운영정원_합본.xlsx")
    final_df.to_excel(output_filename, index=False)
    print(f"\nSaved merged file to: {output_filename}")

if __name__ == "__main__":
    main()
