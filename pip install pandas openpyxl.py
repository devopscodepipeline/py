import pandas as pd

# Load the organized sheet
organized_df = pd.read_excel('organized_sheet.xlsx')

# Load the unorganized sheet
unorganized_df = pd.read_excel('unorganized_sheet.xlsx')

# Ensure all columns in the organized sheet are present in the unorganized sheet
# Fill missing columns with NaN, and reorder to match the organized sheet
rearranged_df = unorganized_df.reindex(columns=organized_df.columns)

# Save the rearranged sheet to a new Excel file
rearranged_df.to_excel('rearranged_sheet.xlsx', index=False)

print("The unorganized sheet has been rearranged and saved as 'rearranged_sheet.xlsx'.")
