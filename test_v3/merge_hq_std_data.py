import pandas as pd
import os

base_dir = r"d:\Workspace\Manage_TO\test_v3"

def get_latest_file(base_dir, keyword):
    files = [f for f in os.listdir(base_dir) if keyword in f and f.endswith(".xlsx")]
    # Avoid picking the result of a previous run or other merged files
    files = [f for f in files if "합본" not in f] 
    files.sort(key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)
    if files:
        return os.path.join(base_dir, files[0])
    return None

def main():
    # 1. Identify Input Files
    std_file = get_latest_file(base_dir, "본부_기준정원_데이터")
    temp_file = get_latest_file(base_dir, "본부_한시정원_데이터")
    
    if not std_file or not temp_file:
        print(f"Error: Could not find input files.\nStd: {std_file}\nTemp: {temp_file}")
        return

    print(f"Loading Standard: {std_file}")
    df_std = pd.read_excel(std_file)
    print(f"Rows: {len(df_std)}")

    print(f"Loading Temporary: {temp_file}")
    df_temp = pd.read_excel(temp_file)
    print(f"Rows: {len(df_temp)}")
    
    # 2. Perform Concatenation (Union)
    # Std file (650 rows) + Temp file (6 rows) = 656 rows.
    
    print("Concatenating files...")
    final_df = pd.concat([df_std, df_temp], ignore_index=True, sort=False)
    
    # 3. Verification
    print(f"Total Rows: {len(final_df)}")
    expected_rows = len(df_std) + len(df_temp)
    if len(final_df) != expected_rows:
        print(f"[WARNING] Row count mismatch! Expected {expected_rows}, got {len(final_df)}")
    
    # Check if new columns exist
    new_cols = ["한시", "한시_존속기한", "한시_업무"]
    for c in new_cols:
        if c not in final_df.columns:
            print(f"[ERROR] Column {c} missing in result.")
    
    # 4. Save
    output_filename = os.path.join(base_dir, "(251001)본부_기준정원_합본.xlsx")
    final_df.to_excel(output_filename, index=False)
    print(f"\nSaved merged file to: {output_filename}")

if __name__ == "__main__":
    main()
