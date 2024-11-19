# Record Late Deliveries (see if there is room for improvement)
# Currently showing top performers. Also show weak performers?
# Use Geopandas?
# 1. Fix monthly -> yearly


import pandas as pd
from sqlalchemy import create_engine
import urllib.parse
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Connection details
server = "johndroescher.com"
default_db = "Fall_2024"
username = "emirdincer34"
password = "CCny24018534"

# Construct the connection string
connection_string = (
    "mssql+pyodbc:///?odbc_connect=" + 
    urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={server};DATABASE={default_db};"
        f"UID={username};PWD={password};"
        "TrustServerCertificate=yes"
    )
)

# Create the SQLAlchemy engine
engine = create_engine(connection_string)

# Function to fetch data
def fetch_data(table_name):
    query = f"SELECT * FROM {table_name}"
    try:
        data = pd.read_sql(query, engine)
        print(f"Loaded {table_name} with {data.shape[0]} rows and {data.shape[1]} columns")
        return data
    except Exception as e:
        print(f"Error loading {table_name}: {e}")

# Fetch data from tables
tables = ["proj_customers", "proj_sales", "proj_products", "proj_exchange_rates", "proj_stores"]
data = {table: fetch_data(table) for table in tables}

# Create a directory to save images
output_dir = "visualizations"
os.makedirs(output_dir, exist_ok=True)

# Extract individual datasets
sales_data = data["proj_sales"]
product_data = data["proj_products"]
customer_data = data["proj_customers"]
exchange_rates_data = data["proj_exchange_rates"]
stores_data = data["proj_stores"]

# Rename columns for consistency
product_data = product_data.rename(columns={"Product_Name": "ProductName"})
customer_data = customer_data.rename(columns={"Name": "CustomerName"})

# Calculate revenue and profitability
product_data['revenue_per_unit'] = product_data['Unit_Price_USD'] - product_data['Unit_Cost_USD']
sales_data = sales_data.merge(product_data[['ProductKey', 'revenue_per_unit', 'ProductName', 'Unit_Price_USD', 'Unit_Cost_USD']], on="ProductKey", how="left")
sales_data['revenue'] = sales_data['Quantity'] * sales_data['revenue_per_unit']

# Visualization 1: Sales Trends Over Time
sales_data['Order_Date'] = pd.to_datetime(sales_data['Order_Date'], errors='coerce')
sales_data['Month'] = sales_data['Order_Date'].dt.to_period('M').dt.to_timestamp()
monthly_sales = sales_data.groupby('Month')['revenue'].sum().reset_index()

plt.figure(figsize=(12, 6))
sns.lineplot(data=monthly_sales, x='Month', y='revenue')
plt.title("Monthly Sales Trends (Revenue in USD)")
plt.xlabel("Month")
plt.ylabel("Revenue (USD)")
plt.xticks(rotation=45)
plt.savefig(f"{output_dir}/monthly_sales_trends.png")
plt.show()

# Visualization 2: Revenue by Product
product_revenue = sales_data.groupby('ProductKey')['revenue'].sum().reset_index()
product_revenue = product_revenue.merge(product_data[['ProductKey', 'ProductName']], on='ProductKey', how='left').sort_values(by='revenue', ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(data=product_revenue.head(10), x='revenue', y='ProductName', palette='viridis')
plt.title("Top 10 Products by Revenue (Revenue in USD)")
plt.xlabel("Revenue (USD)")
plt.ylabel("Product")
plt.savefig(f"{output_dir}/top_products_by_revenue.png")
plt.show()

# Visualization 3: Revenue by Store (by State)
store_revenue = sales_data.groupby('StoreKey')['revenue'].sum().reset_index()
store_revenue = store_revenue.merge(stores_data[['StoreKey', 'State']], on='StoreKey', how='left').sort_values(by='revenue', ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(data=store_revenue.head(10), x='revenue', y='State', palette='plasma')
plt.title("Top 10 States by Store Revenue (Revenue in USD)")
plt.xlabel("Revenue (USD)")
plt.ylabel("State")
plt.savefig(f"{output_dir}/top_states_by_revenue.png")
plt.show()

# Visualization 4: Customer Lifetime Value
customer_revenue = sales_data.groupby('CustomerKey')['revenue'].sum().reset_index()
customer_revenue = customer_revenue.merge(customer_data[['CustomerKey', 'CustomerName']], on='CustomerKey', how='left').sort_values(by='revenue', ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(data=customer_revenue.head(10), x='revenue', y='CustomerName', palette='coolwarm')
plt.title("Top 10 Customers by Lifetime Revenue (Revenue in USD)")
plt.xlabel("Revenue (USD)")
plt.ylabel("Customer")
plt.savefig(f"{output_dir}/top_customers_by_revenue.png")
plt.show()

# Visualization 5: Revenue vs Store Size
stores_data['StoreSize'] = stores_data['Square_Meters']
store_revenue_size = store_revenue.merge(stores_data[['StoreKey', 'StoreSize']], on='StoreKey', how='left')

plt.figure(figsize=(12, 6))
sns.scatterplot(data=store_revenue_size, x='StoreSize', y='revenue', size='revenue', alpha=0.7)
plt.title("Revenue vs Store Size (Revenue in USD)")
plt.xlabel("Store Size (Square Meters)")
plt.ylabel("Revenue (USD)")
plt.savefig(f"{output_dir}/revenue_vs_store_size.png")
plt.show()

# Visualization 6: Profit Marvin vs Sales Volume
sales_data['Profit_Margin_per_Unit'] = sales_data['Unit_Price_USD'] - sales_data['Unit_Cost_USD']

plt.figure(figsize=(12, 6))
sns.scatterplot(data=sales_data, x='Profit_Margin_per_Unit', y='Quantity', alpha=0.7)
plt.title("Impact of Profit Margin on Sales Volume")
plt.xlabel("Profit Margin per Unit (USD)")
plt.ylabel("Quantity Sold")
plt.grid(visible=True)
plt.tight_layout()
plt.savefig(f"{output_dir}/profit_margin_vs_sales_volume.png")
plt.show()

# Visualization 7: Adjusted Revenue by Exchange Rates
exchange_rates_data['Date'] = pd.to_datetime(exchange_rates_data['Date'])
adjusted_sales = pd.merge(sales_data, exchange_rates_data, left_on='Order_Date', right_on='Date', how='left')
adjusted_sales['AdjustedRevenue'] = adjusted_sales['revenue'] * adjusted_sales['Exchange']

plt.figure(figsize=(12, 6))
sns.lineplot(data=adjusted_sales, x='Order_Date', y='AdjustedRevenue', hue='Currency_Code')
plt.title("Adjusted Revenue Trends by Exchange Rates (Revenue in USD)")
plt.xlabel("Date")
plt.ylabel("Adjusted Revenue (USD)")
plt.savefig(f"{output_dir}/adjusted_revenue_trends.png")
plt.show()

# Visualization 8: Regional Sales Performance
regional_sales = sales_data.merge(stores_data[['StoreKey', 'State', 'Country']], on='StoreKey', how='left')
region_revenue = regional_sales.groupby(['State', 'Country'])['revenue'].sum().reset_index()
region_revenue = region_revenue.sort_values(by='revenue', ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(data=region_revenue.head(10), x='revenue', y='State', palette='coolwarm')
plt.title("Top 10 Regions by Revenue")
plt.xlabel("Revenue (USD)")
plt.ylabel("State")
plt.savefig(f"{output_dir}/top_regions_by_revenue.png")
plt.show()

# Visualization 9: Sales Efficiency
store_efficiency = store_revenue_size.copy()
store_efficiency['Efficiency'] = store_efficiency['revenue'] / store_efficiency['StoreSize']

plt.figure(figsize=(12, 6))
sns.barplot(data=store_efficiency.sort_values(by='Efficiency', ascending=False).head(10), x='Efficiency', y='State', palette='Reds')
plt.title("Top 10 Most Efficient Stores")
plt.xlabel("Revenue per Square Meter")
plt.ylabel("State")
plt.savefig(f"{output_dir}/store_efficiency.png")
plt.show()

# Visualization 10: Top Customers by Orders
customer_orders = sales_data.groupby('CustomerKey')['Quantity'].sum().reset_index()
customer_orders = customer_orders.merge(customer_data[['CustomerKey', 'CustomerName']], on='CustomerKey', how='left').sort_values(by='Quantity', ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(data=customer_orders.head(10), x='Quantity', y='CustomerName', palette='Blues')
plt.title("Top 10 Customers by Total Orders")
plt.xlabel("Total Orders")
plt.ylabel("Customer")
plt.savefig(f"{output_dir}/top_customers_by_orders.png")
plt.show()

# Visualization 11: Product Profitability Analysis
profitability = product_data.copy()
profitability['TotalProfit'] = profitability['revenue_per_unit'] * sales_data.groupby('ProductKey')['Quantity'].sum().reset_index()['Quantity']
profitability = profitability.sort_values(by='TotalProfit', ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(data=profitability.head(10), x='TotalProfit', y='ProductName', palette='Purples')
plt.title("Top 10 Products by Profitability")
plt.xlabel("Total Profit (USD)")
plt.ylabel("Product")
plt.savefig(f"{output_dir}/product_profitability.png")
plt.show()

# Visualization 12: Market Basket Analysis
sales_data['ProductPair'] = sales_data['ProductName'] + " & " + sales_data['Order_Number'].astype(str)
pair_counts = sales_data.groupby('ProductPair').size().reset_index(name='Counts').sort_values(by='Counts', ascending=False)

plt.figure(figsize=(12, 6))
sns.barplot(data=pair_counts.head(10), x='Counts', y='ProductPair', palette='Greens')
plt.title("Top 10 Product Pairs")
plt.xlabel("Number of Orders")
plt.ylabel("Product Pair")
plt.savefig(f"{output_dir}/market_basket.png")
plt.show()

# Visualization 13: Product Categories
category_revenue = sales_data.merge(product_data[['ProductKey', 'Category']], on='ProductKey', how='left')
category_revenue_summary = category_revenue.groupby('Category')['revenue'].sum().reset_index()

plt.figure(figsize=(12, 6))
sns.barplot(data=category_revenue_summary, x='revenue', y='Category', palette='Oranges')
plt.title("Revenue by Product Category")
plt.xlabel("Revenue (USD)")
plt.ylabel("Category")
plt.savefig(f"{output_dir}/product_category_revenue.png")
plt.show()

# Visualization 14: Customer Demographics vs Sales
customer_data['Age'] = datetime.now().year - pd.to_datetime(customer_data['Birthday']).dt.year
customer_age_revenue = sales_data.merge(customer_data[['CustomerKey', 'Age']], on='CustomerKey', how='left')
customer_age_revenue_summary = customer_age_revenue.groupby('Age')['revenue'].sum().reset_index()

plt.figure(figsize=(12, 6))
sns.lineplot(data=customer_age_revenue_summary, x='Age', y='revenue', marker='o')
plt.title("Revenue by Customer Age")
plt.xlabel("Age")
plt.ylabel("Revenue (USD)")
plt.savefig(f"{output_dir}/customer_age_revenue.png")
plt.show()

# Visualization 15: Average Order Value Over Time
average_order_value = sales_data.groupby('Order_Date')['revenue'].mean().reset_index()

plt.figure(figsize=(12, 6))
sns.lineplot(data=average_order_value, x='Order_Date', y='revenue', marker='o')
plt.title("Average Order Value Over Time")
plt.xlabel("Order Date")
plt.ylabel("Average Revenue per Order (USD)")
plt.savefig(f"{output_dir}/average_order_value.png")
plt.show()
