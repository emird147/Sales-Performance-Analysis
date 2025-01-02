import pandas as pd
from sqlalchemy import create_engine
import urllib.parse
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os

# Connect to SQL
server = "server"
default_db = "db"
username = "username"
password = "password"

connection_string = (
    "mssql+pyodbc:///?odbc_connect=" + 
    urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 18 for SQL Server}};"
        f"SERVER={server};DATABASE={default_db};"
        f"UID={username};PWD={password};"
        "TrustServerCertificate=yes"
    )
)

# SQL Engine
engine = create_engine(connection_string)

# Fetch Data
def fetch_data(table_name, columns="*", where_clause=None):
    query = f"SELECT {columns} FROM {table_name}"
    if where_clause:
        query += f" WHERE {where_clause}"
    try:
        data = pd.read_sql(query, engine)
        print(f"Loaded {table_name} with {data.shape[0]} rows and {data.shape[1]} columns")
        return data
    except Exception as e:
        print(f"Error loading {table_name}: {e}")

# Filter Data
tables = {
    "proj_customers": "CustomerKey, Name, Birthday",
    "proj_sales": "Order_Date, Quantity, ProductKey, StoreKey, CustomerKey, Order_Number",
    "proj_products": "ProductKey, Product_Name, Unit_Price_USD, Unit_Cost_USD, Category",
    "proj_exchange_rates": "Date, Exchange, Currency", 
    "proj_stores": "StoreKey, State, Country, Square_Meters"
}

data = {table: fetch_data(table, columns=columns) for table, columns in tables.items()}

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
stores_data = stores_data.rename(columns={"Square_Meters": "StoreSize"})

# Calculate revenue and profitability
product_data['revenue_per_unit'] = product_data['Unit_Price_USD'] - product_data['Unit_Cost_USD']
sales_data = sales_data.merge(product_data[['ProductKey', 'revenue_per_unit', 'ProductName', 'Unit_Price_USD', 'Unit_Cost_USD']], on="ProductKey", how="left")
sales_data['Revenue'] = sales_data['Quantity'] * sales_data['Unit_Price_USD']
sales_data['Profit'] = sales_data['Quantity'] * sales_data['revenue_per_unit']

# Investigate Monthly Sales Trends
sales_data['Order_Date'] = pd.to_datetime(sales_data['Order_Date'], errors='coerce')
sales_data['Month'] = sales_data['Order_Date'].dt.to_period('M').dt.to_timestamp()
monthly_sales = sales_data.groupby('Month')[['Revenue', 'Profit']].sum().reset_index()

# Highlight High and Low Sales Periods for Revenue and Profit
high_sales_period_revenue = monthly_sales.loc[monthly_sales['Revenue'].idxmax()]
low_sales_period_revenue = monthly_sales.loc[monthly_sales['Revenue'].idxmin()]

high_sales_period_profit = monthly_sales.loc[monthly_sales['Profit'].idxmax()]
low_sales_period_profit = monthly_sales.loc[monthly_sales['Profit'].idxmin()]

# Visualize Monthly Sales Trends
plt.figure(figsize=(14, 8))
sns.lineplot(data=monthly_sales, x='Month', y='Revenue', label='Revenue', color='blue', linewidth=2)
sns.lineplot(data=monthly_sales, x='Month', y='Profit', label='Profit', color='green', linewidth=2)

# Annotate high and low sales periods
plt.annotate(f"High Revenue: {high_sales_period_revenue['Month'].strftime('%Y-%m')}\n${high_sales_period_revenue['Revenue']:,.2f}", 
             xy=(high_sales_period_revenue['Month'], high_sales_period_revenue['Revenue']),
             xytext=(high_sales_period_revenue['Month'], high_sales_period_revenue['Revenue'] + 500000),
             arrowprops=dict(facecolor='black', arrowstyle="->"), fontsize=10)

plt.annotate(f"Low Revenue: {low_sales_period_revenue['Month'].strftime('%Y-%m')}\n${low_sales_period_revenue['Revenue']:,.2f}", 
             xy=(low_sales_period_revenue['Month'], low_sales_period_revenue['Revenue']),
             xytext=(low_sales_period_revenue['Month'] - pd.DateOffset(months=2), low_sales_period_revenue['Revenue'] - 1500000),
             arrowprops=dict(facecolor='red', arrowstyle="->"), fontsize=10)

plt.annotate(f"High Profit: {high_sales_period_profit['Month'].strftime('%Y-%m')}\n${high_sales_period_profit['Profit']:,.2f}", 
             xy=(high_sales_period_profit['Month'], high_sales_period_profit['Profit']),
             xytext=(high_sales_period_profit['Month'] + pd.DateOffset(months=5), high_sales_period_profit['Profit'] + 700000),
             arrowprops=dict(facecolor='blue', arrowstyle="->"), fontsize=10)

plt.annotate(f"Low Profit: {low_sales_period_profit['Month'].strftime('%Y-%m')}\n${low_sales_period_profit['Profit']:,.2f}", 
             xy=(low_sales_period_profit['Month'], low_sales_period_profit['Profit']),
             xytext=(low_sales_period_profit['Month'] - pd.DateOffset(months=5), low_sales_period_profit['Profit'] - 900000),
             arrowprops=dict(facecolor='orange', arrowstyle="->"), fontsize=10)

plt.title("Monthly Sales and Profit Trends with High/Low Periods")
plt.xlabel("Year")
plt.ylabel("Amount (USD)")
plt.legend()
plt.grid(True)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(f"{output_dir}/monthly_sales_and_profit_trends.png")
plt.show()

# Regional Sales Performance
regional_sales = sales_data.merge(stores_data[['StoreKey', 'State', 'Country']], on='StoreKey', how='left')
region_revenue = regional_sales.groupby(['State', 'Country'])['Profit'].sum().reset_index()
region_revenue = region_revenue.sort_values(by='Profit', ascending=False)

# Visualize Regional Sales Performance
plt.figure(figsize=(12, 6))
sns.barplot(data=region_revenue.head(10), x='Profit', y='State', palette='coolwarm')
plt.title("Top 10 Regions by Profit")
plt.xlabel("Profit (USD)")
plt.ylabel("State")
plt.savefig(f"{output_dir}/top_regions_by_profit.png")
plt.show()

# Dip Period Analysis
monthly_sales['Change'] = monthly_sales['Revenue'].diff()
dip_periods = monthly_sales[monthly_sales['Change'] < 0].sort_values(by='Change').head(10)

# Visualize Dip Periods
plt.figure(figsize=(12, 6))
sns.barplot(data=dip_periods, x='Month', y='Change', palette='Reds_d')
plt.title("Top Dip Periods in Revenue")
plt.xlabel("Date")
plt.ylabel("Revenue Change (USD)")
plt.xticks(rotation=45)
plt.grid(True)
plt.tight_layout()
plt.savefig(f"{output_dir}/dip_periods_revenue.png")
plt.show()

# Define and recalculate top_products_quantity
top_products_quantity = (
    sales_data.groupby("ProductName")
    .agg({"Quantity": "sum", "Profit": "sum"})
    .sort_values(by="Quantity", ascending=False)
    .head(5)
    .reset_index()
)

# Define and recalculate top_products_profitability
top_products_profitability = (
    sales_data.groupby("ProductName")
    .agg({"Profit": "sum", "Quantity": "sum"})
    .sort_values(by="Profit", ascending=False)
    .head(5)
    .reset_index()
)

# Sort legend for quantity by descending order
sales_data_top_quantity = sales_data[sales_data['ProductName'].isin(top_products_quantity['ProductName'])]
sales_trends_quantity = sales_data_top_quantity.groupby(['Month', 'ProductName']).agg({"Quantity": "sum"}).reset_index()
sales_trends_quantity['ProductName'] = pd.Categorical(
    sales_trends_quantity['ProductName'],
    categories=top_products_quantity['ProductName'],
    ordered=True
)

plt.figure(figsize=(14, 8))
sns.lineplot(data=sales_trends_quantity, x='Month', y='Quantity', hue='ProductName', marker='o', hue_order=top_products_quantity['ProductName'])
plt.title("Top Products by Quantity Sold Over Time")
plt.xlabel("Year")
plt.ylabel("Quantity Sold")
plt.legend(title="Product Name", loc='upper left', bbox_to_anchor=(1, 1))
plt.grid(True)
plt.tight_layout()
plt.savefig(f"{output_dir}/top_products_quantity_over_time.png")
plt.show()

# Print Top Products by Quantity Sold with Profit
print("Top Products by Quantity Sold:")
print(top_products_quantity[['ProductName', 'Quantity', 'Profit']])

# Sort legend for profitability by descending order
sales_data_top_profit = sales_data[sales_data['ProductName'].isin(top_products_profitability['ProductName'])]
sales_trends_profit = sales_data_top_profit.groupby(['Month', 'ProductName']).agg({"Profit": "sum"}).reset_index()
sales_trends_profit['ProductName'] = pd.Categorical(
    sales_trends_profit['ProductName'],
    categories=top_products_profitability['ProductName'],
    ordered=True
)

plt.figure(figsize=(14, 8))
sns.lineplot(data=sales_trends_profit, x='Month', y='Profit', hue='ProductName', marker='o', palette='cool', hue_order=top_products_profitability['ProductName'])
plt.title("Top Products by Profitability Over Time")
plt.xlabel("Year")
plt.ylabel("Profit (USD)")
plt.legend(title="Product Name", loc='upper left', bbox_to_anchor=(1, 1))
plt.grid(True)
plt.tight_layout()
plt.savefig(f"{output_dir}/top_products_profitability_over_time.png")
plt.show()

# Print Top Products by Profitability with Quantity
print("Top Products by Profitability:")
print(top_products_profitability[['ProductName', 'Profit', 'Quantity']])

# Add Year and Quarter columns
sales_data['Order_Date'] = pd.to_datetime(sales_data['Order_Date'], errors='coerce')
sales_data['Year'] = sales_data['Order_Date'].dt.year
sales_data['Quarter'] = sales_data['Order_Date'].dt.quarter

# Filter Q1 and Q4 data
q1_sales = sales_data[sales_data['Quarter'] == 1]
q4_sales = sales_data[sales_data['Quarter'] == 4]

# Calculate average Unit Price for Q1 and Q4
avg_unit_price_q1 = q1_sales.groupby('ProductName')['Unit_Price_USD'].mean().rename('Avg_Unit_Price_Q1')
avg_unit_price_q4 = q4_sales.groupby('ProductName')['Unit_Price_USD'].mean().rename('Avg_Unit_Price_Q4')

# Top 3 products by profitability for Q1 and Q4 for each year
q1_top_products_by_year = (
    q1_sales.groupby(['Year', 'ProductName'])
    .agg({'Profit': 'sum'})
    .sort_values(by=['Year', 'Profit'], ascending=[True, False])
    .groupby('Year')
    .head(3)
    .reset_index()
)

q4_top_products_by_year = (
    q4_sales.groupby(['Year', 'ProductName'])
    .agg({'Profit': 'sum'})
    .sort_values(by=['Year', 'Profit'], ascending=[True, False])
    .groupby('Year')
    .head(3)
    .reset_index()
)

# Merge Q1 and Q4 prices into comparison_data
comparison_data = pd.concat([q1_top_products_by_year.assign(Quarter='Q1'), q4_top_products_by_year.assign(Quarter='Q4')])
comparison_data = comparison_data.merge(avg_unit_price_q1, on='ProductName', how='left')
comparison_data = comparison_data.merge(avg_unit_price_q4, on='ProductName', how='left')

# Determine if Q4 price is lower than Q1 price for discounts
comparison_data['Discounted'] = comparison_data['Avg_Unit_Price_Q4'] < comparison_data['Avg_Unit_Price_Q1']

# Plot the bar chart for each year
years = sales_data['Year'].dropna().unique()
for year in years:
    plt.figure(figsize=(12, 8))
    yearly_data = comparison_data[comparison_data['Year'] == year]

    # Create a barplot with Q1 and Q4 colors
    ax = sns.barplot(
        data=yearly_data,
        x='ProductName',
        y='Profit',
        hue='Quarter',
        dodge=True,
        palette={'Q1': 'blue', 'Q4': 'orange'}
    )

    plt.title(f'Top 3 Most Profitable Products: Q1 vs Q4 in {int(year)}')
    plt.xlabel('Product Name')
    plt.ylabel('Profit (USD)')
    plt.legend(title='Quarter', loc='upper right')
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y')

    # Annotate bars with discount info
    for bar, (_, row) in zip(ax.patches, yearly_data.iterrows()):
        discount_text = 'Discount' if row['Discounted'] else 'No Discount'
        plt.text(
            bar.get_x() + bar.get_width() / 2, 
            bar.get_height(), 
            discount_text,
            ha='center', va='bottom', fontsize=9, color='black'
        )

    plt.tight_layout()
    plt.savefig(f"{output_dir}/top_products_profit_q1_vs_q4_{int(year)}_discounts.png")
    plt.show()

    # Filter sales data for 2016-2021
sales_data_filtered = sales_data[(sales_data['Year'] >= 2016) & (sales_data['Year'] <= 2021)]

# Top 3 most bought products in Q1 vs Q4 for each year
def top_3_products_by_quantity(quarter):
    top_products = sales_data_filtered[sales_data_filtered['Quarter'] == quarter]
    top_products = top_products.groupby(['Year', 'ProductName'])['Quantity'].sum().reset_index()
    top_products = top_products.sort_values(by=['Year', 'Quantity'], ascending=[True, False])
    top_products = top_products.groupby('Year').head(3)
    return top_products

q1_top_products = top_3_products_by_quantity(1)
q4_top_products = top_3_products_by_quantity(4)

# Merge Q1 and Q4 for comparison
def plot_top_products_q1_q4(q1_data, q4_data):
    years = q1_data['Year'].unique()
    for year in years:
        q1_year_data = q1_data[q1_data['Year'] == year]
        q4_year_data = q4_data[q4_data['Year'] == year]
        
        fig, ax1 = plt.subplots(figsize=(12, 8))
        sns.barplot(
            data=q1_year_data,
            x='ProductName',
            y='Quantity',
            color='blue',
            label='Q1'
        )
        sns.barplot(
            data=q4_year_data,
            x='ProductName',
            y='Quantity',
            color='orange',
            label='Q4'
        )
        
        plt.title(f'Top 3 Most Bought Products: Q1 vs Q4 in {year}')
        plt.xlabel('Product Name')
        plt.ylabel('Quantity Sold')
        plt.xticks(rotation=45, ha='right')
        plt.legend()
        plt.tight_layout()
        plt.savefig(f"{output_dir}/top_3_products_q1_vs_q4_{year}.png")
        plt.show()

plot_top_products_q1_q4(q1_top_products, q4_top_products)

# Top 5 products by profit per year
top_5_products_profit = (
    sales_data.groupby(['Year', 'ProductName'])
    .agg({'Profit': 'sum'})
    .sort_values(by=['Year', 'Profit'], ascending=[True, False])
    .groupby('Year')
    .head(5)
    .reset_index()
)

# Top 5 products by quantity per year
top_5_products_quantity = (
    sales_data.groupby(['Year', 'ProductName'])
    .agg({'Quantity': 'sum'})
    .sort_values(by=['Year', 'Quantity'], ascending=[True, False])
    .groupby('Year')
    .head(5)
    .reset_index()
)

# Plot top 5 products by profit per year
years = sales_data['Year'].dropna().unique()
for year in years:
    plt.figure(figsize=(12, 8))
    yearly_profit_data = top_5_products_profit[top_5_products_profit['Year'] == year]
    sns.barplot(
        data=yearly_profit_data,
        x='ProductName',
        y='Profit',
        palette='Blues_d'
    )
    plt.title(f'Top 5 Products by Profit in {int(year)}')
    plt.xlabel('Product Name')
    plt.ylabel('Profit (USD)')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(f"{output_dir}/top_5_products_profit_{int(year)}.png")
    plt.show()

    # Plot top 5 products by quantity per year
for year in years:
    plt.figure(figsize=(12, 8))
    yearly_quantity_data = top_5_products_quantity[top_5_products_quantity['Year'] == year]
    sns.barplot(
        data=yearly_quantity_data,
        x='ProductName',
        y='Quantity',
        palette='Greens_d'
    )
    plt.title(f'Top 5 Products by Quantity in {int(year)}')
    plt.xlabel('Product Name')
    plt.ylabel('Quantity Sold')
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(f"{output_dir}/top_5_products_quantity_{int(year)}.png")
    plt.show()

    # Print top 5 products by profit per year
print("Top 5 Products by Profit Per Year:")
print(top_5_products_profit)

# Print top 5 products by quantity per year
print("Top 5 Products by Quantity Per Year:")
print(top_5_products_quantity)

# Calculate state performance
state_performance = sales_data.merge(stores_data, on='StoreKey').groupby('State').agg({
    'Profit': 'sum',
    'StoreSize': 'sum'
}).reset_index()
state_performance['Efficiency'] = state_performance['Profit'] / state_performance['StoreSize']

# Top 5 states by profit
top_5_states_profit = state_performance.nlargest(5, 'Profit')

# Top 5 states by profit
top_5_states_profit = state_performance.nlargest(5, 'Profit')

# Plot top 5 states by profit
plt.figure(figsize=(12, 8))
sns.barplot(
    data=top_5_states_profit,
    x='State',
    y='Profit',
    palette='Purples_d'
)
plt.title('Top 5 States by Profit')
plt.xlabel('State')
plt.ylabel('Profit (USD)')
plt.tight_layout()
plt.savefig(f"{output_dir}/top_5_states_profit.png")
plt.show()

# Top 5 efficient states
top_5_states_efficiency = state_performance.nlargest(5, 'Efficiency')

# Plot top 5 efficient states
plt.figure(figsize=(12, 8))
sns.barplot(
    data=top_5_states_efficiency,
    x='State',
    y='Efficiency',
    palette='Oranges_d'
)
plt.title('Top 5 Efficient States by Profit per Square Meter')
plt.xlabel('State')
plt.ylabel('Efficiency (Profit/Square Meter)')
plt.tight_layout()
plt.savefig(f"{output_dir}/top_5_states_efficiency.png")
plt.show()

# Combine top 5 profitable and efficient states
top_states_combined = pd.concat([top_5_states_profit, top_5_states_efficiency]).drop_duplicates(subset=['State'])

# Plot combined chart of top states by profit and efficiency
plt.figure(figsize=(12, 8))
top_states_combined['Metric'] = ['Profit' if state in top_5_states_profit['State'].values else 'Efficiency' for state in top_states_combined['State']]
sns.barplot(
    data=top_states_combined,
    x='State',
    y='Profit',
    hue='Metric',
    palette={'Profit': 'purple', 'Efficiency': 'orange'}
)
plt.title('Combined Top States by Profit and Efficiency')
plt.xlabel('State')
plt.ylabel('Value')
plt.legend(title='Metric')
plt.tight_layout()
plt.savefig(f"{output_dir}/combined_top_states.png")
plt.show()

# Top 5 products by profit and quantity for top states
top_products_top_states = (
    sales_data.merge(stores_data, on='StoreKey')
    .merge(top_states_combined[['State']], on='State')
    .groupby(['State', 'ProductName'])
    .agg({'Profit': 'sum', 'Quantity': 'sum'})
    .reset_index()
)

# Plot top 5 products by profit and quantity for each state
for state in top_products_top_states['State'].unique():
    state_data = top_products_top_states[top_products_top_states['State'] == state]

    # Top 5 by both Profit and Quantity
    top_5_combined = state_data.nlargest(5, 'Profit')
    plt.figure(figsize=(12, 8))
    ax = sns.barplot(
        data=top_5_combined,
        x='ProductName',
        y='Profit',
        color='blue',
        label='Profit'
    )
    ax2 = ax.twinx()
    sns.lineplot(
        data=top_5_combined,
        x='ProductName',
        y='Quantity',
        color='green',
        marker='o',
        ax=ax2,
        label='Quantity'
    )
    ax.set_ylabel('Profit (USD)')
    ax2.set_ylabel('Quantity Sold')
    plt.title(f'Top 5 Products by Profit and Quantity in {state}')
    ax.set_xlabel('Product Name')
    ax.set_xticklabels(top_5_combined['ProductName'], rotation=45, ha='right')
    ax.legend(loc='upper left')
    ax2.legend(loc='upper right')
    plt.tight_layout()
    plt.savefig(f"{output_dir}/top_5_products_combined_{state}.png")
    plt.show()

    # Top 5 efficient states
top_5_states_efficiency = state_performance.nlargest(5, 'Efficiency')

# Combine top 5 profitable and efficient states
top_states_combined = pd.concat([top_5_states_profit, top_5_states_efficiency]).drop_duplicates(subset=['State'])

# Plot combined chart of top states by profit and efficiency
plt.figure(figsize=(12, 8))
top_states_combined['Metric'] = ['Profit' if state in top_5_states_profit['State'].values else 'Efficiency' for state in top_states_combined['State']]
sns.barplot(
    data=top_states_combined,
    x='State',
    y='Profit',
    hue='Metric',
    palette={'Profit': 'purple', 'Efficiency': 'orange'}
)
plt.title('Combined Top States by Profit and Efficiency')
plt.xlabel('State')
plt.ylabel('Value')
plt.legend(title='Metric')
plt.tight_layout()
plt.savefig(f"{output_dir}/combined_top_states.png")
plt.show()

# Top 5 products by profit and quantity for top states
top_products_top_states = (
    sales_data.merge(stores_data, on='StoreKey')
    .merge(top_states_combined[['State']], on='State')
    .groupby(['State', 'ProductName'])
    .agg({'Profit': 'sum', 'Quantity': 'sum'})
    .reset_index()
)

# Plot top 5 products by profit and quantity for each state
for state in top_products_top_states['State'].unique():
    state_data = top_products_top_states[top_products_top_states['State'] == state]

    # Top 5 by both Profit and Quantity
    top_5_combined = state_data.nlargest(5, 'Profit')
    plt.figure(figsize=(12, 8))
    ax = sns.barplot(
        data=top_5_combined,
        x='ProductName',
        y='Profit',
        color='blue',
        label='Profit'
    )
    ax2 = ax.twinx()
    sns.lineplot(
        data=top_5_combined,
        x='ProductName',
        y='Quantity',
        color='green',
        marker='o',
        ax=ax2,
        label='Quantity'
    )
    ax.set_ylabel('Profit (USD)')
    ax2.set_ylabel('Quantity Sold')
    plt.title(f'Top 5 Products by Profit and Quantity in {state}')
    ax.set_xlabel('Product Name')
    ax.set_xticklabels(top_5_combined['ProductName'], rotation=45, ha='right')
    ax.legend(loc='upper left')
    ax2.legend(loc='upper right')
    plt.tight_layout()
    plt.savefig(f"{output_dir}/top_5_products_combined_{state}.png")