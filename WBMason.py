import pandas as pd

# Define the data as a dictionary
data = {
    'Year': ['2021', '2021', '2021', '2021', '2021', '2022', '2022', '2022', '2022', '2022', '2022', '2022', '2022', '2022', '2022', '2023', '2023', '2023', '2023'],
    'Month': ['Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Jan', 'Feb', 'Mar', 'Apr'],
    'Count': [60, 91, 71, 123, 149, 108, 96, 136, 82, 85, 69, 64, 96, 92, 111, 66, 64, 107, 23]
}

# Create a DataFrame from the dictionary
df = pd.DataFrame(data)

# Calculate total counts by year
total_counts = df.groupby('Year')['Count'].sum().reset_index()
total_counts.rename(columns={'Count': 'Total Count'}, inplace=True)

# Create a DataFrame for the grand total
grand_total = pd.DataFrame({'Year': ['Grand Total'], 'Total Count': [df['Count'].sum()]})

# Concatenate total counts and grand total
summary_df = pd.concat([total_counts, grand_total], ignore_index=True)

# Save to Excel
with pd.ExcelWriter('WB_Mason_Summary.xlsx') as writer:
    df.to_excel(writer, sheet_name='Monthly Data', index=False)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

print("WB Mason data has been processed and saved to WB_Mason_Summary.xlsx")
