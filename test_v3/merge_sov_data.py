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

def merge_and_save(df1, df2, output_name):
    print(f"Concatenating {len(df1)} + {len(df2)} rows...")
    final_df = pd.concat([df1, df2], ignore_index=True, sort=False)
    
    print(f"Total Rows: {len(final_df)}")
    expected = len(df1) + len(df2)
    if len(final_df) != expected:
        print(f"[WARNING] Row count mismatch! Expected {expected}, got {len(final_df)}")
        
    output_path = os.path.join(base_dir, output_name)
    final_df.to_excel(output_path, index=False)
    print(f"Saved to: {output_path}\n")

def main():
    # Identify Files
    std_file = get_latest_file(base_dir, "소속기관_기준정원_데이터")
    ops_file = get_latest_file(base_dir, "소속기관_운영정원_데이터")
    temp_file = get_latest_file(base_dir, "소속기관_한시정원_데이터")
    
    if not (std_file and ops_file and temp_file):
        print("Error: Could not find one or more input files.")
        print(f"Std: {std_file}")
        print(f"Ops: {ops_file}")
        print(f"Temp: {temp_file}")
        return

    print("Loading files...")
    df_std = pd.read_excel(std_file)
    print(f"Standard Loaded: {std_file} ({len(df_std)} rows)")
    
    df_ops = pd.read_excel(ops_file)
    print(f"Operational Loaded: {ops_file} ({len(df_ops)} rows)")
    
    df_temp = pd.read_excel(temp_file)
    print(f"Temporary Loaded: {temp_file} ({len(df_temp)} rows)")
    print("-" * 30)
    
    # 1. Merge Standard + Temp
    print("1. Merging Standard + Temp -> 소속기관_기준정원_합본.xlsx")
    merge_and_save(df_std, df_temp, "(251001)소속기관_기준정원_합본.xlsx")
    
    # 2. Merge Operational + Temp
    print("2. Merging Operational + Temp -> 소속기관_운영정원_합본.xlsx")
    merge_and_save(df_ops, df_temp, "(251001)소속기관_운영정원_합본.xlsx")

if __name__ == "__main__":
    main()
