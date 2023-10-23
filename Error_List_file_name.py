import pandas as pd

#List down column
# Read the Excel file into a DataFrame
excel_file = r"D:\PYDATAANALYSIS\Error_List_file_name copy.xlsx"
df = pd.read_excel(excel_file)

# List to store the concatenated values
concatenated_values = []

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
            concatenated_values.append(last_part)

# Create a DataFrame with the concatenated values
concatenated_df = pd.DataFrame(concatenated_values, columns=["Concatenated"])

# Export the DataFrame to another file
exported_excel_file = "exported_concatenated_values.xlsx"
concatenated_df.to_excel(exported_excel_file, index=False)

print(f"Concatenated values exported to {exported_excel_file}")


#ROW
# Read the Excel file into a DataFrame
excel_file = r"D:\PYDATAANALYSIS\exported_concatenated_values.xlsx"

# Create a new DataFrame with the formatted string as the only row
new_row = pd.DataFrame({'Formatted': [df['Concatenated'].str.cat(sep=' OR ')]})

# Define the sheet name for the new row
sheet_name = 'NewSheet'

# Open the Excel file and write the new row to the specified sheet
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a') as writer:
    new_row.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Formatted value added to sheet '{sheet_name}' in {excel_file}")

