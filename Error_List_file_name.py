import pandas as pd

# Read the Excel file into a DataFrame
excel_file = r"D:\PYDATAANALYSIS\Error_List_file_name copy.xlsx"
df = pd.read_excel(excel_file)

# List to store the extracted last parts
last_parts = []

# Process and check all rows in the DataFrame
for index, row in df.iterrows():
    # Access and check data in each row
    # For example, you can check a specific column by column name
    column_value = row['Error']
    
    # Define your condition here, for example, check if column_value starts with "Unable to find file"
    if column_value.startswith("Unable to find file"):
        # Extract the last part of the file name by splitting using "\"
        parts = column_value.split("\\")
        if len(parts) > 0:
            last_part = parts[-1]
            last_parts.append(last_part)

# Create a DataFrame with the extracted last parts
last_parts_df = pd.DataFrame(last_parts, columns=["Last Part"])

# Export the DataFrame to another file
exported_excel_file = "exported_last_parts.xlsx"
last_parts_df.to_excel(exported_excel_file, index=False)

print(f"Last parts exported to {exported_excel_file}")
