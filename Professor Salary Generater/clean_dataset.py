import pandas as pd
import numpy as np # Used for handling NaN values

# --- Configuration ---
# Define the path to your Excel file.
# IMPORTANT: Ensure this path is correct and accessible on your system.
excel_file_path = r"C:\Users\ACER\Downloads\PAY BILL-CLASS 1 & 2-JUNE-2025.xlsx"

# Define output file paths for the cleaned data
cleaned_csv_output_path = r"C:\Users\ACER\Downloads\PAY BILL-CLASS 1 & 2-JUNE-2025_cleaned.csv"
cleaned_excel_output_path = r"C:\Users\ACER\Downloads\PAY BILL-CLASS 1 & 2-JUNE-2025_cleaned.xlsx"

# --- Step 1: Load Data Correctly by Skipping Initial Metadata and Setting Proper Header ---
try:
    # Load the Excel file:
    # - skiprows=[0]: This skips the very first row (Excel's row 1) which contains
    #                 the general polytechnic title.
    # - header=0: After skipping row 0, the new row 0 (which was Excel's row 2,
    #             containing "Sr. No." and "Section of establishment...") will be
    #             used as the header.
    df = pd.read_excel(excel_file_path, skiprows=[0], header=0)

    print("--- Initial Load Preview (after header adjustment) ---")
    print(df.head(10))
    print("\n")

except FileNotFoundError:
    print(f"Error: The file '{excel_file_path}' was not found.")
    print("Please ensure the file path is correct and the file exists at that location.")
    exit() # Exit the script if the file is not found
except Exception as e:
    print(f"An error occurred while reading the Excel file: {e}")
    exit() # Exit if any other error occurs during loading

# --- Step 2: Consolidate Multi-Row Records and Clean Column Names ---

# Rename the key columns for clarity and easier manipulation
df = df.rename(columns={
    'Sr. No.': 'Sr_No',
    'Section of establishment & Name of incumbent': 'Employee_Details'
})

# Drop columns that are completely empty (all NaN values). This removes most 'Unnamed' columns.
df_cleaned_cols = df.dropna(axis=1, how='all')

# Initialize a list to store the cleaned, consolidated records
consolidated_records = []
current_employee_record = {}

# Iterate through each row of the DataFrame to consolidate
for index, row in df_cleaned_cols.iterrows():
    # Check if this row marks the beginning of a new employee record.
    # A new record starts when 'Sr_No' is a number (not NaN) and is not the old header value '1'.
    # This specifically targets rows like 'SHRI P. R. DAVE' that have a 'Sr_No'.
    if pd.notna(row['Sr_No']) and isinstance(row['Sr_No'], (int, float)) and row['Sr_No'] not in [1.0, '1']:
        # If we have a partially built record from the previous iteration, save it
        if current_employee_record:
            consolidated_records.append(current_employee_record)

        # Start a new record with initial values
        current_employee_record = {
            'Sr_No': int(row['Sr_No']) if pd.notna(row['Sr_No']) else None,
            'Employee_Name': str(row['Employee_Details']).strip() if pd.notna(row['Employee_Details']) else '',
            'Designation': '',
            'Pay_Scale_Raw': '', # Temporarily store raw pay scale string
            # Add other relevant columns here if you identify more from your sheet
            # For example, if 'Unnamed: 2' in your raw data contained a 'Department':
            # 'Department': str(row.get('Unnamed: 2', '')).strip() if pd.notna(row.get('Unnamed: 2')) else ''
        }
    elif current_employee_record: # This row is a continuation of the current employee's record
        details = str(row['Employee_Details']).strip() if pd.notna(row['Employee_Details']) else ''

        # Check for keywords to extract information
        if 'PRINCIPAL' in details.upper() or 'PROFESSOR' in details.upper() or 'LECTURER' in details.upper() or 'HOD' in details.upper() or 'ACCOUNTANT' in details.upper():
            current_employee_record['Designation'] = details
        elif 'PAY SCALE' in details.upper():
            current_employee_record['Pay_Scale_Raw'] = details
        # Add more 'elif' conditions here for other specific patterns you observe in your data
        # that should be extracted into new columns.
        # Example:
        # elif 'DATE OF JOINING' in details.upper():
        #    current_employee_record['Joining_Date'] = details # You'd then parse this into a date object

# Add the very last processed record after the loop finishes
if current_employee_record:
    consolidated_records.append(current_employee_record)

# Create a new DataFrame from the consolidated records
df_final_cleaned = pd.DataFrame(consolidated_records)

print("--- Preview after Consolidating Multi-Row Records ---")
print(df_final_cleaned.head(10))
print("\n")

# --- Step 3: Refine Data Types and Extract Structured Information (e.g., Pay Scale) ---

# Clean 'Employee_Name': Remove common prefixes like 'SHRI', 'SMT', 'MS', 'MR'
df_final_cleaned['Employee_Name'] = df_final_cleaned['Employee_Name'].astype(str).str.replace(r'^(SHRI|SMT|MS|MR)\s+', '', regex=True).str.strip()

# Extract Min and Max Pay Scale from 'Pay_Scale_Raw'
def parse_pay_scale(pay_scale_str):
    if pd.isna(pay_scale_str) or not isinstance(pay_scale_str, str):
        return np.nan, np.nan
    # Find numbers that look like ranges (e.g., 144200-218200)
    matches = pd.Series(pay_scale_str).str.extract(r'(\d+)\s*-\s*(\d+)').iloc[0]
    if not matches.empty and pd.notna(matches[0]) and pd.notna(matches[1]):
        try:
            return int(matches[0]), int(matches[1])
        except ValueError:
            return np.nan, np.nan
    return np.nan, np.nan

# Apply the function to create new columns
df_final_cleaned[['Min_Pay_Scale', 'Max_Pay_Scale']] = df_final_cleaned['Pay_Scale_Raw'].apply(lambda x: pd.Series(parse_pay_scale(x)))

# Drop the original raw pay scale column as it's now parsed
df_final_cleaned = df_final_cleaned.drop(columns=['Pay_Scale_Raw'])

# Ensure 'Sr_No' is an integer type, allowing for potential NaNs
df_final_cleaned['Sr_No'] = pd.to_numeric(df_final_cleaned['Sr_No'], errors='coerce').astype('Int64')

# Fill any remaining empty strings in 'Designation' or 'Employee_Name' with NaN for consistency
df_final_cleaned['Designation'] = df_final_cleaned['Designation'].replace('', np.nan)
df_final_cleaned['Employee_Name'] = df_final_cleaned['Employee_Name'].replace('', np.nan)


print("--- Final Cleaned DataFrame Preview ---")
print(df_final_cleaned.head(10))
print("\n")

print("--- Final Cleaned DataFrame Info ---")
df_final_cleaned.info()
print("\n")

print("--- Final Cleaned Basic Statistics (numeric columns) ---")
print(df_final_cleaned.describe())
print("\n")

print("--- Final Cleaned Basic Statistics (categorical columns) ---")
print(df_final_cleaned.describe(include='object'))
print("\n")


# --- Step 4: Save the Cleaned Dataset ---

try:
    # Save to CSV
    df_final_cleaned.to_csv(cleaned_csv_output_path, index=False, encoding='utf-8')
    print(f"Cleaned dataset saved successfully to CSV: {cleaned_csv_output_path}")

    # Save to Excel
    df_final_cleaned.to_excel(cleaned_excel_output_path, index=False)
    print(f"Cleaned dataset saved successfully to Excel: {cleaned_excel_output_path}")

except Exception as e:
    print(f"An error occurred while saving the cleaned dataset: {e}")