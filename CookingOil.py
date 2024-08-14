import pandas as pd

# Create a new DataFrame for the CookingOil sheet
columns = ['Ref', 'Category', 'Month', 'Year', 'Date', 'Building', 'Tonnage', 'Waste Type', 'Vendor', 'Cost', 'Avoided cost', 'Waste Stream', 'Weight(lbs)']
data = []

# List of months to loop through
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

# Initialize some constants
category = "Non C&D"
year = 2023
waste_type = "Cooking Oil"
vendor = "American ByProduct"
cost = "-"
waste_stream = "Diverted"
msw_cost = 95  # MSW cost per ton

# Loop through months twice, first for BID EAST, then for BID WEST
for i in range(2):  # Two cycles for two buildings (BID EAST, BID WEST)
    building = "BID EAST" if i == 0 else "BID WEST"
    for month in months:
        # Calculate some fields
        date = f"{month}-23"
        
        # Calculate the tonnage and avoided cost; these are placeholders, replace them with actual formulas
        # Using 0.0 for now for Tonnage and Weight(lbs)
        weight_lbs = 0.0  # This should be replaced with the correct formula
        tonnage = weight_lbs / 2000
        avoided_cost = tonnage * msw_cost

        # Add row to the data list
        data.append([
            "conc Month Building Waste Type Vendor",  # Ref (this can be modified)
            category,
            month,
            year,
            date,
            building,
            tonnage,
            waste_type,
            vendor,
            cost,
            avoided_cost,
            waste_stream,
            weight_lbs
        ])

# Create the DataFrame from the data
cooking_oil_df = pd.DataFrame(data, columns=columns)

# Load the existing Excel file
with pd.ExcelWriter("Master 2023.xlsx", mode='a', engine='openpyxl') as writer:
    # Write the new sheet into the existing file
    cooking_oil_df.to_excel(writer, sheet_name="CookingOil", index=False)
