
# Step 1: Load the Dataset
import pandas as pd
import plotly.express as px

# Read data from Excel file
data = pd.read_excel('Ecommerce Data.xlsx')

# Step 2: Initial Exploration
# View first and last few rows
print(data.head(2))
print(data.tail())

# Check shape and column types
print("Dataset shape:", data.shape)
print("Data types:\n", data.dtypes)
print("\nData summary:")
data.info()
print(data.describe())

# Step 3: Data Cleaning

# 3.1 Check for missing values
print("\nMissing values in each column:\n", data.isnull().sum())

# 3.2 Check for duplicate rows
duplicates = data.duplicated()
print("Number of duplicate records:", duplicates.sum())

# If any, display duplicated rows
if duplicates.sum() > 0:
    print(data[duplicates])

# 3.3 Drop rows with null values in essential columns
columns_to_clean = ['Customer Gender', 'Location', 'Zone', 'Delivery Type', 'Product Category', 'SubCategory']
for col in columns_to_clean:
    data = data[data[col].notna()]

# Verify cleaning
print("\nNulls after cleaning:\n", data.isnull().sum())
data.info()

# Step 4: Insights & Analysis

# 4.1 Top product categories by frequency
print("\nProduct category distribution:\n", data['Product Category'].value_counts())

# 4.2 Most frequently sold products
print("\nTop 5 selling products:\n", data['Product'].value_counts().head(5))
print("Total unique products:", data['Product'].nunique())

# 4.3 Average rating per category
category_ratings = data.pivot_table(index='Product Category', values='Rating', aggfunc='mean').sort_values(by='Rating', ascending=False)
print("\nAverage Rating by Category:\n", category_ratings)

# 4.4 Create rating categories
rating_bins = [0, 1, 2, 3, 4, 5]
rating_labels = ['Very Dissatisfied', 'Dissatisfied', 'Neutral', 'Satisfied', 'Very Satisfied']
data['Rating_Group'] = pd.cut(data['Rating'], bins=rating_bins, labels=rating_labels)

# Preview categorized ratings
print(data[['Rating', 'Rating_Group']].head())

# 4.5 Reorder pivot with margins (Overall Avg)
rating_summary = data.pivot_table(index='Product Category', values='Rating', aggfunc='mean', margins=True).sort_values(by='Rating', ascending=False)
ordered_index = [cat for cat in rating_summary.index if cat != 'All'] + ['All']
rating_summary = rating_summary.reindex(ordered_index)
print("\nDetailed Rating Summary:\n", rating_summary)

# 4.6 Regional Sales Analysis
location_sales = data.groupby('Location')['Sale Price'].sum().sort_values(ascending=False).head(10)
print("\nTop 10 Locations by Total Sales:\n", location_sales)

# Step 5: Visualization

# 5.1 Category-wise Order Quantity
order_counts = data.pivot_table(index='Product Category', values='Order Quantity', aggfunc='count').sort_values(by='Order Quantity', ascending=False)

# 5.2 Pie Chart using Plotly
fig = px.pie(
    values=order_counts['Order Quantity'],
    names=order_counts.index,
    title='Order Distribution by Product Category'
)
fig.update_layout(title_x=0.5)
fig.show()

# Final display of pivot table
print("\nOrder Quantity by Product Category:\n", order_counts)
