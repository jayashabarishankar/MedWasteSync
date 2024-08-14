import pandas as pd
import numpy as np
import math

# Load data from the Excel file
data_dump = pd.read_excel('DATA DUMP.xls', sheet_name='Summary', header=None)

# Define the column names
column_names = [
    "Ref", "Category", "Month", "Year", "Building", 
    "Tonnage", "Waste Type", "Vendor", "Cost", "Avoided Cost", "Waste Stream"
]

# Create the original DataFrame with the specified column names
df_original = pd.DataFrame(columns=column_names)

# Create a lookup table for the Building assignments
lookup_table = {
    (8058523, '001'): 'BID EAST',
    (8112041, '002'): 'BID EAST',
    (8057291, '001'): 'BID WEST',
    (8112041, '001'): 'BID WEST',
    (8057291, '002'): 'BID WEST',
    (8112041, '008'): 'BID WEST',
    (8064949, '001'): 'RN',
    (8112041, '003'): 'RN'
}

# Month lookup table
month_lookup = {
    '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr',
    '05': 'May', '06': 'Jun', '07': 'Jul', '08': 'Aug',
    '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
}

reverse_month_lookup = {v: k for k, v in month_lookup.items()}

# Fill the original DataFrame with initial rows from the data_dump
for i in range(5, len(data_dump)):
    customer_number = str(data_dump.at[i, 1]).strip()
    site_number = str(data_dump.at[i, 2]).strip().zfill(3)
    value = str(data_dump.at[i, 4]).strip()
    ref_string = str(data_dump.at[i, 3]).strip()
    month_digits = value[-2:]
    month = month_lookup.get(month_digits, 'Unknown')
    building_name = lookup_table.get((int(customer_number), site_number), 'Unknown')
    ref = f"{customer_number}{site_number} {ref_string}"
    base_row_data = {
        "Ref": ref,
        "Category": "Non C & D",
        "Month": month,
        "Year": "2023",
        "Building": building_name,
        "Tonnage": "-",
        "Vendor": "Stericycle",
        "Avoided Cost": "-",
        "Waste Stream": "Regulated Waste"
    }
    df_original = pd.concat([df_original, pd.DataFrame([base_row_data])], ignore_index=True)

#df_original = df_original.drop_duplicates()

# Make a copy of the original DataFrame to create DataFrame 2 with same columns
df2 = pd.DataFrame(columns=column_names)

# Loop through each row in df_original and data_dump to add rows to df2
for i in range(len(df_original)):
    base_row_data = df_original.iloc[i].to_dict()

    # Check corresponding row in data_dump
    data_dump_index = i + 5
    waste_types = {
        7: "RMW",
        10: "Sharps",
        13: "Path/ Chemo"
    }

    rows_to_add = []

    for col_idx, waste_type in waste_types.items():
        cost_value = data_dump.at[data_dump_index, col_idx]
        weight_value = data_dump.at[data_dump_index, col_idx-1]
        if pd.notna(cost_value) and isinstance(cost_value, (int, float)):
            new_row = base_row_data.copy()
            new_row["Waste Type"] = waste_type
            new_row["Cost"] = cost_value
            new_row["Weight"] = weight_value
            rows_to_add.append(new_row)

    # Add the base row and then any additional rows for waste types
    df2 = pd.concat([df2, pd.DataFrame([base_row_data])], ignore_index=True)
    if rows_to_add:
        df2 = pd.concat([df2, pd.DataFrame(rows_to_add)], ignore_index=True)

# Calculate Tonnage
df2["Tonnage"] = df2["Weight"] 

# Drop unnecessary rows and columns
#f2 = df2.drop(df2.index[::4])
df2 = df2.drop(columns=['Weight'])

def custom_round(x):
    if pd.isna(x):  # Check if x is NaN
        return x  # Return NaN unchanged
    if isinstance(x, (int, float)):  # Check if x is an int or float
        return math.ceil(x) if x - int(x) >= 0.5 else round(x)  # Round up if .5 or more
    return x  # Return the value unchanged if it's not a number


# Apply custom rounding function
#df2 = df2.applymap(custom_round)

df2 = df2.dropna(subset=['Tonnage', 'Cost'], how='all')




def consolidate_rows(df):
    # Group by the specified columns and sum the Tonnage and Cost columns
    grouped_df = df.groupby(['Month', 'Year', 'Building', 'Waste Type'], as_index=False).agg({
        'Tonnage': 'sum',#aggregate sum and then divide by 2000
        'Cost': 'sum',
        'Ref': 'first',  # Use 'first' to keep the first occurrence's values for non-aggregated columns
        'Category': 'first',
        'Vendor': 'first',
        'Avoided Cost': 'first',
        'Waste Stream': 'first'
    })
   

    return grouped_df




#df2 = df2[df2['Month'] != 'Sep']




# Use the function to consolidate df2
df2 = consolidate_rows(df2)

df2['Tonnage'] = df2['Tonnage'] / 2000

# Display the consolidated DataFrame

# Save to Excel if needed

def round_numeric_cells(df, decimals=2):
    # Apply rounding to numeric cells in the entire DataFrame
    for column in df.columns:
        if pd.api.types.is_numeric_dtype(df[column]):
            df[column] = df[column].apply(lambda x: round(x, decimals) if pd.notna(x) else x)
    return df

# Apply the function
df_rounded = round_numeric_cells(df2)




# Save the sorted DataFrame to Excel
df_rounded.to_excel('Output_Final.xlsx', index=False)
