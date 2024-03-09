import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Read the Excel file into a pandas DataFrame
df0 = pd.read_excel('Global Bike Sales.xlsx')

# Create a copy of the original DataFrame to work with
df1 = df0.copy()

# Display the first few rows of the DataFrame
df1.head()

# Question 1: Calculate the overall revenue for all years and identify the year with the highest revenue.
overall_rev = df1['Revenue'].sum()
print("Qn 1: Overall revenue is $", '{:,}'.format(overall_rev))

highest_rev = df1.groupby('Calendar Year').sum()['Revenue'].sort_values(ascending=False).head(1)
print("Highest revenue is $", '{:,}'.format(highest_rev.iloc[0]), "in", highest_rev.index[0])

# Question 2: Calculate the revenue for E-Bike in 2018
ebike_rev_2018 = df1[(df1['Material Desc'] == 'E-Bike Tailwind') & (df1['Calendar Year'] == 2018)]['Revenue'].sum()
print("Qn 2: Revenue for E-Bike in 2018 is $", '{:,}'.format(round(ebike_rev_2018, 2)))

# Question 3: Calculate the revenue for E-Bike in 2017
ebike_rev_2017 = df1[(df1['Material Desc'] == 'E-Bike Tailwind') & (df1['Calendar Year'] == 2017)]['Revenue'].sum()
print("Qn 3: Revenue for E-Bike in 2017 is $", '{:,}'.format(round(ebike_rev_2017, 2)))

# Question 4: Calculate the sales quantity of T-shirts for Airport Bikes in 2015
sales_qtty_tshirts_airport_2015 = df1[(df1['Material Desc'] == 'T-shirt') & 
                                       (df1['Customer Desc'] == 'Airport Bikes') & 
                                       (df1['Calendar Year'] == 2015)]['Sales Quantity'].sum()
print('Qn 4: Sales quantity for T-shirts for Airport Bikes in 2015 is', '{:,}'.format(sales_qtty_tshirts_airport_2015))

# Question 5: Identify the year with the highest sales quantity
highest_sales_year = df1.groupby('Calendar Year').sum()['Sales Quantity'].idxmax()
print("Qn 5: Year with the highest sales quantity is", highest_sales_year)

# Question 6: Identify the year with the lowest revenue
lowest_revenue_year = df1.groupby('Calendar Year').sum()['Revenue'].idxmin()
print("Qn 6: Year with the lowest revenue is", lowest_revenue_year)

# Question 7: Identify the material with the lowest revenue overall
material_lowest_rev = df1.groupby(['Material', 'Material Desc']).sum()['Revenue'].idxmin()[1]
print('Qn 7: Material with lowest revenue is', material_lowest_rev)

# Question 8: Identify the division with the highest net sales in the year with the highest net sales
year_highest_net_sales = highest_rev.index[0]
division_highest_net_sales = df1[df1['Calendar Year'] == year_highest_net_sales].groupby('Division').sum()['Net Sales'].idxmax()
print(f"Qn 8: In the year {year_highest_net_sales}, the {division_highest_net_sales} Division had the highest net sales.")

# Question 9: Identify the customer with the highest net sales in the division and year from Question 8
customer_highest_net_sales = df1[(df1['Calendar Year'] == year_highest_net_sales) & 
                                 (df1['Division'] == division_highest_net_sales)].groupby('Customer Desc').sum()['Net Sales'].idxmax()
print(f"Qn 9: For the {division_highest_net_sales} Division in {year_highest_net_sales}, the customer with the highest net sales is {customer_highest_net_sales}.")

# Question 10: Identify the division with the highest net sales in the year with the lowest net sales
year_lowest_net_sales = df1.groupby('Calendar Year').sum()['Net Sales'].idxmin()
division_highest_net_sales_lowest_year = df1[df1['Calendar Year'] == year_lowest_net_sales].groupby('Division').sum()['Net Sales'].idxmax()
print(f"Qn 10: In the year {year_lowest_net_sales}, the {division_highest_net_sales_lowest_year} Division had the highest net sales.")

# Question 11: Identify the customer with the highest net sales in the division from Question 10 for all years
customer_highest_net_sales_div_lowest_year = df1[df1['Division'] == division_highest_net_sales_lowest_year].groupby('Customer Desc').sum()['Net Sales'].idxmax()
print(f"Qn 11: For the {division_highest_net_sales_lowest_year} Division, the customer with the highest net sales for all years is {customer_highest_net_sales_div_lowest_year}.")

# Question 12: Identify the customer that provided the highest net sales for Accessories (Division AS) in 2017
customer_highest_net_sales_accessories_2017 = df1[(df1['Division'] == 'AS') & 
                                                  (df1['Calendar Year'] == 2017)].groupby('Customer Desc').sum()['Net Sales'].idxmax()
print(f"Qn 12: In 2017, for Accessories (Division AS), the customer with the highest net sales is {customer_highest_net_sales_accessories_2017}.")

# Question 13: Calculate the total revenue for the top three customers for Bicycles (Division BI) in 2019
df1_BI_2019 = df1[(df1['Division'] == 'BI') & (df1['Calendar Year'] == 2019)]
top_three_customers_BI_2019 = df1_BI_2019.groupby('Customer Desc').sum()['Revenue'].nlargest(3)
total_revenue_top_three_customers_BI_2019 = top_three_customers_BI_2019.sum()
print("Qn 13: Total revenue for the top three customers for Bicycles (Division BI) in 2019 is $", '{:,}'.format(total_revenue_top_three_customers_BI_2019))

# Question 14: Calculate the average revenue per customer
avg_revenue_per_customer = df1.groupby('Customer Desc').sum()['Revenue'].mean()
print("Qn 14: Average revenue per customer is $", '{:,}'.format(avg_revenue_per_customer))

# Question 15: Calculate the average revenue per customer each year
avg_revenue_per_customer_each_year = df1.groupby(['Calendar Year', 'Customer Desc']).sum()['Revenue'].groupby('Calendar Year').mean()
print("Qn 15: Average revenue per customer each year:")
print(avg_revenue_per_customer_each_year)

# Question 16: Calculate the average revenue per customer per year per division
avg_revenue_per_customer_per_year_per_division = df1.groupby(['Division', 'Calendar Year', 'Customer Desc']).sum()['Revenue'].groupby(['Division', 'Calendar Year']).mean()
print("Qn 16: Average revenue per customer per year per division:")
print(avg_revenue_per_customer_per_year_per_division)

# Question 17: Identify the year with the highest average revenue per customer and the amount
year_highest_avg_revenue_per_customer = avg_revenue_per_customer_each_year.idxmax()
amount_highest_avg_revenue_per_customer = avg_revenue_per_customer_each_year.max()
print("Qn 17: Year with the highest average revenue per customer:", year_highest_avg_revenue_per_customer)
print("Amount of the highest average revenue per customer: $", '{:,}'.format(amount_highest_avg_revenue_per_customer))

# Question 18: Identify the customer with the highest average revenue during the year
customer_highest_avg_revenue_per_year = df1.groupby(['Customer Desc', 'Calendar Year']).sum()['Revenue'].groupby('Calendar Year').idxmax()
print("18: Customer with the highest average revenue during each year:")
print(customer_highest_avg_revenue_per_year)

# Question 19: Analyze revenue trends for the US and Germany (DE)
revenue_trends_by_country = df1.groupby(['Calendar Year', 'Country']).sum()['Revenue'].reset_index()
sns.lineplot(data=revenue_trends_by_country, x='Calendar Year', y='Revenue', hue='Country')
plt.title('Qn 19: Revenue Trends by Country')
plt.show()

# Question 20: Analyze overall revenue trends
overall_revenue_trends = df1.groupby('Calendar Year').sum()['Revenue'].reset_index()
sns.lineplot(data=overall_revenue_trends, x='Calendar Year', y='Revenue')
plt.title('Qn 20: Overall Revenue Trends')
plt.show()

# Question 21: Analyze seasonality in revenue
revenue_seasonality = df1.groupby('Calendar Month').sum()['Revenue'].reset_index()
sns.lineplot(data=revenue_seasonality, x='Calendar Month', y='Revenue')
plt.title('Qn 21: Seasonality in Revenue')
plt.show()

# Question 22: Analyze revenue trends for different products
product_categories = df1['Product Category'].unique()
fig, axes = plt.subplots(nrows=3, ncols=2, figsize=(15, 10))
for i, category in enumerate(product_categories):
    row = i // 2
    col = i % 2
    ax = axes[row, col]
    category_data = df1[df1['Product Category'] == category]
    sns.lineplot(data=category_data, x='Calendar Year', y='Revenue', hue='Country', ax=ax)
    ax.set_title(f'Qn 22: Revenue Trend for {category}')
plt.tight_layout()
plt.show()

# Question 23: Identify materials without significant seasonality
materials = df1['Material'].unique()
fig, axes = plt.subplots(nrows=2, ncols=2, figsize=(15, 10))
for i in range(4):
    row = i // 2
    col = i % 2
    ax = axes[row, col]
    material_subset = materials[i*7:(i+1)*7]
    material_data = df1[df1['Material'].isin(material_subset)]
    sns.lineplot(data=material_data, x='Calendar Month', y='Revenue', hue='Material', ax=ax)
    ax.set_title(f'Qn 23: Revenue Trend for Material Category {i+1}')
plt.tight_layout()
plt.show()

# Question 24: Analyze the percentage contribution of each customer to revenue over the years
customer_contribution_percentage = df1.groupby('Customer Desc').sum()['Revenue'] / df1['Revenue'].sum() * 100
customer_contribution_percentage = customer_contribution_percentage.sort_values(ascending=False)
print("Qn 24: Customer with the highest percentage contribution to revenue:", customer_contribution_percentage.index[0])

# Question 25: Identify the year with the highest overall gross margin
df1['Gross Profit Margin'] = df1['Net Sales'] - df1['Cost of Goods Sold']
year_highest_gross_margin = df1.groupby('Calendar Year').sum()['Gross Profit Margin'].idxmax()
amount_highest_gross_margin = df1.groupby('Calendar Year').sum()['Gross Profit Margin'].max()
print("Qn 25: Year with the highest overall gross margin:", year_highest_gross_margin)
print("Amount of the highest overall gross margin: $", '{:,}'.format(amount_highest_gross_margin))

# Question 26: Calculate the gross margin as a percentage of sales for Germany and the US in the year with the highest gross margin
gross_margin_DE_highest_year = df1[(df1['Country'] == 'DE') & (df1['Calendar Year'] == year_highest_gross_margin)]['Gross Profit Margin'].sum()
gross_margin_US_highest_year = df1[(df1['Country'] == 'US') & (df1['Calendar Year'] == year_highest_gross_margin)]['Gross Profit Margin'].sum()
total_sales_highest_year = df1[df1['Calendar Year'] == year_highest_gross_margin]['Net Sales'].sum()
gross_margin_percentage_DE = (gross_margin_DE_highest_year / total_sales_highest_year) * 100
gross_margin_percentage_US = (gross_margin_US_highest_year / total_sales_highest_year) * 100
print("Qn 26 a) Gross margin as a percentage of sales for Germany in the year with the highest gross margin:", round(gross_margin_percentage_DE, 2), "%")
print("Qn 26 b) Gross margin as a percentage of sales for Germany in the year with the highest gross margin:", round(gross_margin_percentage_US, 2), "%")
