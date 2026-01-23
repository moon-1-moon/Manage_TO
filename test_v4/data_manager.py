import pandas as pd
import os

class DataManager:
    def __init__(self, base_dir=None):
        if base_dir is None:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))
        else:
            self.base_dir = base_dir
            
        self.files = {
            "STD": "(251001)기준정원.xlsx",
            "OPS": "(251001)운영정원.xlsx"
        }
        self.dataframes = {}
        self.current_type = "STD" # or "OPS"
        
        # Columns used for slicing/filtering
        self.filter_cols = [
            "부서1", "부서2", "부서3", "부서4", "부서5", "부서6",
            "직군", "직급", "직렬"
        ]
        
    def load_data(self):
        """Loads both Excel files into memory."""
        for key, filename in self.files.items():
            path = os.path.join(self.base_dir, filename)
            if os.path.exists(path):
                print(f"Loading {filename}...")
                try:
                    df = pd.read_excel(path)
                    # Ensure numeric columns are numeric
                    if '정원' in df.columns:
                        df['정원'] = pd.to_numeric(df['정원'], errors='coerce').fillna(0)
                except Exception as e:
                    print(f"Error loading {filename}: {e}")
                    df = pd.DataFrame(columns=self.filter_cols + ["정원"])
                self.dataframes[key] = df
            else:
                print(f"File not found: {path}")
                self.dataframes[key] = pd.DataFrame(columns=self.filter_cols + ["정원"])

    def get_current_df(self):
        return self.dataframes.get(self.current_type, pd.DataFrame())

    def set_current_type(self, type_key):
        if type_key in self.files:
            self.current_type = type_key

    def get_filtered_data(self, filters):
        """
        Returns a filtered DataFrame.
        filters: dict {col_name: [list of selected values]}
        """
        df = self.get_current_df()
        if df.empty:
            return df
            
        mask = pd.Series([True] * len(df))
        
        for col, selected_vals in filters.items():
            if col in df.columns and selected_vals:
                mask &= df[col].isin(selected_vals)
                
        return df[mask]

    def get_unique_values(self, target_col, current_filters):
        """
        Returns unique values for target_col, respecting OTHER filters.
        This enables 'dynamic' slicer behavior (cascading filters).
        """
        df = self.get_current_df()
        if df.empty or target_col not in df.columns:
            return []
            
        # Apply filters for ALL columns EXCEPT the target_col
        # This allows the user to see other options within the current context,
        # or we could strictly apply all filters. 
        # Usually, for a slicer A, we want to see options compatible with slicers B, C...
        
        mask = pd.Series([True] * len(df))
        for col, selected_vals in current_filters.items():
            if col != target_col and col in df.columns and selected_vals:
                mask &= df[col].isin(selected_vals)
                
        subset = df[mask]
        values = subset[target_col].dropna().unique().tolist()
        return sorted(map(str, values)) # Return as sorted strings

    def add_row(self, row_data, index=0):
        """
        row_data: dict suitable for DataFrame
        index: insertion index (0-based)
        """
        df = self.get_current_df()
        new_row = pd.DataFrame([row_data])
        
        if df.empty:
            self.dataframes[self.current_type] = new_row
        else:
            # Split and concatenate to insert
            if index < 0: index = 0
            if index > len(df): index = len(df)
            
            top = df.iloc[:index]
            bottom = df.iloc[index:]
            self.dataframes[self.current_type] = pd.concat([top, new_row, bottom], ignore_index=True)

    def update_row(self, index, col_name, new_value):
        df = self.get_current_df()
        if 0 <= index < len(df) and col_name in df.columns:
            # Handle numeric conversion for '정원'
            if col_name == '정원':
                try:
                    new_value = float(new_value)
                except:
                    pass
            df.at[index, col_name] = new_value

    def delete_rows(self, indices):
        """indices: list of integers"""
        df = self.get_current_df()
        if not df.empty:
            df.drop(indices, inplace=True)
            df.reset_index(drop=True, inplace=True)
            self.dataframes[self.current_type] = df

    def export_data(self, filepath):
        df = self.get_current_df()
        try:
            df.to_excel(filepath, index=False)
            return True, "Success"
        except Exception as e:
            return False, str(e)

    def save_current_file(self):
        """Overwrites the source file with current data."""
        filename = self.files[self.current_type]
        path = os.path.join(self.base_dir, filename)
        return self.export_data(path)
