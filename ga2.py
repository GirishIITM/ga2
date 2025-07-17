import pandas as pd

xls = pd.ExcelFile("data.xlsx", engine="openpyxl")

# An E-Commerce Platform - has one of its Mother Distribution Centres (DC) set up in
# Mumbai which coordinates with its Three Child DCs: 3. Pune, 2. Aurangabad, 3. Nashik.
# For this assignment, we are considering 12 categories; 5 products each leading up to 50
# products.
# There are 6 data sheets in the Excel File provided to you.
# SKU Master: It contains information of BU, SKU, Brand, Model and Average Price
# Sales Data: It contains information on Date, SKU, City and Volume
# OPN STK: Stock that is available as on 03-01-2025
# Stock Transfer: Contains information on incoming stocks to Pune, Aurangabad and
# Nasik from Mumbai from 03-01-2025 – 31-01-2025
sku_df = pd.read_excel(xls, xls.sheet_names[0])
sales_df = pd.read_excel(xls, xls.sheet_names[1])
stocks_df = pd.read_excel(xls, xls.sheet_names[2])

stocks_df["Category"] = stocks_df["Category"].str.strip()
stocks_df["SKU"] = stocks_df["SKU"].str.strip()

transactions_raw = pd.read_excel(xls, xls.sheet_names[3])

cities = []
dates = []
current_city = ""
for i in range(1, len(transactions_raw.columns)):
    col_name = transactions_raw.columns[i]
    if "Pune" in str(col_name) or col_name == "Pune":
        current_city = "Pune"
    elif "Aurangabad" in str(col_name) or col_name == "Aurangabad":
        current_city = "Aurangabad"
    elif "Nasik" in str(col_name) or col_name == "Nasik":
        current_city = "Nasik"
    else:
        pass

    date_val = transactions_raw.iloc[0, i]
    cities.append(current_city)
    dates.append(date_val)

new_columns = ["SKU"]
for i in range(len(cities)):
    new_columns.append(f"{cities[i]}_{dates[i]}")

transactions_clean = transactions_raw.iloc[1:].copy()
transactions_clean.columns = new_columns

transactions_list = []
for col in transactions_clean.columns[1:]:
    city, date = col.split("_", 1)
    temp_df = transactions_clean[["SKU", col]].copy()
    temp_df["City"] = city
    temp_df["Date"] = pd.to_datetime(date)
    temp_df["Units"] = temp_df[col]
    temp_df = temp_df[["SKU", "City", "Date", "Units"]]
    transactions_list.append(temp_df)

transactions_df = pd.concat(transactions_list, ignore_index=True)
transactions_df = transactions_df.dropna()

transactions_df["SKU"] = transactions_df["SKU"].str.strip()
transactions_with_category = pd.merge(
    sku_df, transactions_df, on="SKU", how="left")
transactions_with_category["Category"] = transactions_with_category[
    "Category"
].str.strip()
transactions_with_category["Date"] = pd.to_datetime(
    transactions_with_category["Date"])
transactions_with_category["Units"] = transactions_with_category["Units"].astype(
    int)
sales_and_sku_df = pd.merge(sales_df, sku_df, on="SKU", how="left")

# 2. For the entire month, what is the total sale value of the game “LTA Wise City”?
# (INTEGER)
sale_value_lta_wise_city = sales_and_sku_df[
    sales_and_sku_df["Product Name"] == "LTA Wise City"
]
sum_sale_value_lta_wise_city = (
    sale_value_lta_wise_city["Price"] * sale_value_lta_wise_city["Sales"]
).sum()
print("Total sale value of LTA Wise City question 2 :",
      sum_sale_value_lta_wise_city)

# 3. What fraction of total sale quantity (Volume) did “Books” category achieve in the first week? (Jan 1 to Jan 7, both days included) (FLOAT between 0 and 1)
# Hint: Construct a Volume Pareto Chart
books_sales_and_sku = sales_and_sku_df["Category"] == "Books"
total_books_sales = sales_and_sku_df[books_sales_and_sku &
                                     (sales_and_sku_df["Date"] >= "2025-01-01") &
                                     (sales_and_sku_df["Date"] <= "2025-01-07")
                                     ]["Sales"].sum()
total_sales = sales_and_sku_df["Sales"].sum()
fraction_books_sales = total_books_sales / total_sales
print(
    "Fraction of total sale quantity for Books category question 3:",
    fraction_books_sales,
)

# 4. What is the maximum sale value by a single SKU in a day across all days?
# (Sale Value = Sale Qty * Price per Qty) (INTEGER)
sales_and_sku_df["Sale Value"] = sales_and_sku_df["Sales"] * \
    sales_and_sku_df["Price"]
date_and_sku_vise_data = sales_and_sku_df.groupby(["SKU", "Date"])[
    "Sale Value"].sum()
single_sku_max_sale_value = date_and_sku_vise_data.max()
print(
    "The maximum sale value by a single SKU in a day across all days question 4:",
    single_sku_max_sale_value,
)

# 5. What is the maximum revenue generating category across all days? (STRING)
date_and_category_vise_df = sales_and_sku_df.groupby(["Category", "Date"])[
    "Sale Value"
].sum()
single_category_max_sale_value = date_and_category_vise_df.max()
single_category_max = date_and_category_vise_df.idxmax()
print(
    "The maximum revenue generating category across all days question 5:",
    single_category_max,
)

# 6. What fraction of total sale value did Mumbai achieve? (across all categories and days)
# (FLOAT between 2 and 1)
total_sales = sales_and_sku_df["Sale Value"].sum()
mumbai_sales = (sales_and_sku_df[sales_and_sku_df["City"] == "Mumbai"])[
    "Sale Value"
].sum()
fraction_mumbai_sales = mumbai_sales / total_sales
print(
    "fraction of total sale value did Mumbai achieve? (across all categories and days) question 6: ",
    fraction_mumbai_sales,
)

# 7. What is the no. of units of household category SKUs are in stock at the end of
# 17th Jan 2025 in Nasik DC? (INTEGER)
household_stocks = stocks_df[stocks_df["Category"] == "Household"]
opening_household_stocks = household_stocks["Nashik"].sum()
household_stocks_transfers_df = transactions_with_category[
    (transactions_with_category["Category"] == "Household")
    & (transactions_with_category["City"] == "Nasik")
    & (transactions_with_category["Date"] <= "2025-01-17")
    & (transactions_with_category["Units"] > 0)
]
household_stocks_transfers = household_stocks_transfers_df["Units"].sum()
total_household_stocks = opening_household_stocks + household_stocks_transfers
print(
    "The no. of units of household category SKUs are in stock at the end of 17th Jan 2025 in Nasik DC question 7:",
    total_household_stocks,
)

# 8. Based on the sales and stock data of Jan 2025, how many average days of
# inventory of SKU M004 are available in Pune? (FLOAT)


def average_days_inventory(sku, city):
    opening_stocks = stocks_df[(stocks_df["SKU"] == sku)][city].values[0]
    incomming_stocks = transactions_df[
        (transactions_df["SKU"] == sku) & (transactions_df["City"] == city)
    ]["Units"].sum()
    sales_average = sales_and_sku_df[
        (sales_and_sku_df["SKU"] == sku) & (sales_and_sku_df["City"] == city)
    ]["Sales"].mean()
    return (opening_stocks + incomming_stocks) / sales_average


average_days_inventory_pune_m004 = average_days_inventory("M004", "Pune")
print(
    "Average days of inventory of SKU M004 in Pune question 8:",
    average_days_inventory_pune_m004,
)

# 9. Which SKU has the highest average days of inventory in Aurangabad? (STRING)
sku_list = stocks_df["SKU"].unique()
def highest_average_days_inventory(city):
    max_days = 0
    max_sku = ""
    for sku in sku_list:
        days = average_days_inventory(sku, city)
        if days > max_days:
            max_days = days
            max_sku = sku
    return max_sku


highest_average_days_inventory_aurangabad = highest_average_days_inventory(
    "Aurangabad")
print(
    "SKU with the highest average days of inventory in Aurangabad question 9:",
    highest_average_days_inventory_aurangabad,
)

# 10. How many SKUs hold at least one weeks’ worth of inventory on average in
# Pune? (INTEGER)
sku_with_week_inventory = 0
for sku in sku_list:
    days_inventory = average_days_inventory(sku, "Pune")
    if days_inventory >= 7:
        sku_with_week_inventory += 1
print(
    "Number of SKUs that hold at least one weeks’ worth of inventory on average in Pune question 10:",
    sku_with_week_inventory,
)

# 11. What is the closing stock of K005 at the end of the month in Nasik? (INTEGER)
closing_stock_k005_nasik = stocks_df[
    (stocks_df["SKU"] == "K005")
]["Nashik"].values[0]
closing_stock_k005_nasik += transactions_df[
    (transactions_df["SKU"] == "K005")
    & (transactions_df["City"] == "Nasik")
]["Units"].sum()
print(
    "Closing stock of K005 at the end of the month in Nasik question 11:",
    closing_stock_k005_nasik,
)