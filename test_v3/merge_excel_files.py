import pandas as pd
import os

base_dir = r"c:\Workspace\Manage_TO\test_v3"

def merge_files(file1, file2, output_file):
    path1 = os.path.join(base_dir, file1)
    path2 = os.path.join(base_dir, file2)
    output_path = os.path.join(base_dir, output_file)

    print(f"Loading {file1}...")
    df1 = pd.read_excel(path1)
    print(f" - Rows: {len(df1)}")

    print(f"Loading {file2}...")
    df2 = pd.read_excel(path2)
    print(f" - Rows: {len(df2)}")

    print("Merging...")
    merged_df = pd.concat([df1, df2], ignore_index=True)
    print(f" - Total Rows: {len(merged_df)}")

    print(f"Saving to {output_file}...")
    merged_df.to_excel(output_path, index=False)
    print("Done.\n")

if __name__ == "__main__":
    # Merge Criterion Quota (기준정원)
    merge_files(
        "(251001)본부_기준정원_합본.xlsx",
        "(251001)소속기관_기준정원_합본.xlsx",
        "(251001)기준정원.xlsx"
    )

    # Merge Operating Quota (운영정원)
    merge_files(
        "(251001)본부_운영정원_합본.xlsx",
        "(251001)소속기관_운영정원_합본.xlsx",
        "(251001)운영정원.xlsx"
    )
