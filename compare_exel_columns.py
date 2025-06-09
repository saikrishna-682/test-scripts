import pandas as pd

def compare_excel_columns(file1_path, file2_path, column_name, sheet_name1=0, sheet_name2=0):
    try:
        # Read the Excel files
        df1 = pd.read_excel(file1_path, sheet_name=sheet_name1)
        df2 = pd.read_excel(file2_path, sheet_name=sheet_name2)
        
        # Check if the column exists in both files
        if column_name not in df1.columns or column_name not in df2.columns:
            print(f"Error: Column '{column_name}' not found in one or both files.")
            return
        
        # Merge dataframes on the specified column to find mismatches
        merged = pd.merge(df1, df2, on=column_name, how='outer', indicator=True)
        
        # Filter rows that are only in one file (mismatches)
        mismatches = merged[merged['_merge'] != 'both']
        
        if mismatches.empty:
            print("No mismatches found in the specified column.")
        else:
            print("Mismatched rows:")
            for index, row in mismatches.iterrows():
                source = 'File 1' if row['_merge'] == 'left_only' else 'File 2'
                print(f"Row from {source}:")
                print(row.drop('_merge'))
                print("-" * 50)
                
    except FileNotFoundError:
        print("Error: One or both files not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage
if __name__ == "__main__":
    file1 = "file1.xlsx"  # Replace with your first Excel file path
    file2 = "file2.xlsx"  # Replace with your second Excel file path
    column_to_compare = "ID"  # Replace with the column name to compare
    
    compare_excel_columns(file1, file2, column_to_compare)
