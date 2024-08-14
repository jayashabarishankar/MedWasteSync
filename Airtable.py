import pandas as pd

# Load the initial sheet (Airtable report)
initial_file = "Airtable report_2023.xlsx"
initial_data = pd.read_excel(initial_file, sheet_name='For master file')

# Create an empty DataFrame for the final output
columns = ['Ref l', 'Ref ll', 'Category', 'Month', 'Year', 'Date', 'Building', 'Tonnage', 'Waste Type', 'Vendor', 'Cost', 'Avoided cost (MSW $95/t)', 'Waste Stream', 'Weight(lbs)']
final_data = pd.DataFrame(columns=columns)

# Loop through each row in the initial data
for index, row in initial_data.iterrows():
    month = row['Month']
    building = row['Campus From']
    waste_type = row['Waste Type']
    weight_lbs = row['Weight(lbs)']

    # Calculate the required fields
    ref_1 = f"{month}{building}{waste_type}"
    ref_2 = f"2023{ref_1}"
    category = "Non C&D"
    year = 2023
    date = f"{month}-23"
    tonnage = weight_lbs / 2000 if weight_lbs else 0
    vendor = "ABC Mover"
    cost = "-"
    avoided_cost = tonnage * 95
    waste_stream = "Landfilled" if waste_type == "Facilities Disposal" else "Diverted"

    # Append the row to the final DataFrame
    final_data = final_data.append({
        'Ref l': ref_1,
        'Ref ll': ref_2,
        'Category': category,
        'Month': month,
        'Year': year,
        'Date': date,
        'Building': building,
        'Tonnage': tonnage,
        'Waste Type': waste_type,
        'Vendor': vendor,
        'Cost': cost,
        'Avoided cost (MSW $95/t)': avoided_cost,
        'Waste Stream': waste_stream,
        'Weight(lbs)': weight_lbs
    }, ignore_index=True)

# Save the final DataFrame to an Excel file
final_file = "Final_Output.xlsx"
final_data.to_excel(final_file, index=False)
