import pandas as pd

# Load the data from the specific Excel sheets (replace with actual sheet names/paths)
reuse_data = pd.read_excel(r'2023 Linen reuse\Data\Linen_Total_2023.xlsx', sheet_name='Data_Pivot')
loss_data = pd.read_excel(r'2023 Linen reuse\Linen loss data\2023\Linen_loss_final.xlsx', sheet_name='Data_Pivot')

# Group by 'Location_updated' and 'Month', then sum the relevant columns
grouped_data = reuse_data.groupby(['Updated_Location', 'Month']).agg({
   # 'G_column': 'sum',  # Replace 'G_column' with the actual column name for the total reuse
    'Sum of Weight(lb)': 'sum',  # Replace 'Weight_column' with the column for weight
    'Sum of Cost($)': 'sum'  # Replace 'Cost_column' with the cost column
}).reset_index()

grouped_data2 = (['Location update', ' ']).agg({
   # 'G_column': 'sum',  # Replace 'G_column' with the actual column name for the total reuse
    'Sum of Weight(lbs)': 'sum',  # Replace 'Weight_column' with the column for weight
    #'Cost_column': 'sum'  # Replace 'Cost_column' with the cost column
}).reset_index()

# Create the additional columns based on the grouped data
grouped_data['Ref l'] = grouped_data.apply(lambda x: f"{x['Month']}BID {x['Updated_Location']}ReuseLinen2023", axis=1)
grouped_data['Ref ll'] = grouped_data.apply(lambda x: f"2023{x['Month']}BID {x['Updated_Location']}Reuse", axis=1)
grouped_data['Category'] = "Non C&D"
grouped_data['Year'] = 2023
grouped_data['Date'] = grouped_data['Month'].astype(str) + "-23"
grouped_data['Building'] = grouped_data['Updated_Location']

# Calculate 'Tonnage' (weight in lbs divided by 2000)
grouped_data['Tonnage'] = grouped_data['Weight(lbs)'] / 2000

# Waste Type, Vendor, and Waste Stream are fixed values
grouped_data['Waste Type'] = "Reuse"
grouped_data['Vendor'] = "Linen"
grouped_data['Waste Stream'] = "Diverted"

# Calculate 'Cost' (based on your second Excel formula)
grouped_data['Cost'] = grouped_data['Sum of Cost($)']  # Assuming Cost_column already contains the calculated value

# Calculate 'Avoided cost (MSW $95/t)' (tonnage * 95)
grouped_data['Avoided cost (MSW $95/t)'] = grouped_data['Tonnage'] * 95

# Calculate 'Weight(lbs)' (based on your first formula)
grouped_data['Weight(lbs)'] = grouped_data2['Sum of Weight(lbs)']- grouped_data['Sum of Weight(lb)'] # This assumes G_column represents the total weight

# Select the columns in the desired order
final_columns = [
    "Ref l", "Ref ll", "Category", "Month", "Year", "Date", 
    "Building", "Tonnage", "Waste Type", "Vendor", "Cost", 
    "Avoided cost (MSW $95/t)", "Waste Stream", "Weight(lbs)"
]
final_df = grouped_data[final_columns]

# Save the final DataFrame to a new Excel file or inspect it
final_df.to_excel('Linen_Final.xlsx', index=False)


