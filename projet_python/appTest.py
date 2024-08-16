#%%
import numpy as nb
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import folium

# %%
file_path = 'sales.xlsx'

# Charger les données
sales_orders = pd.read_excel('Sales.xlsx', sheet_name='Sales Orders')
customers = pd.read_excel('Sales.xlsx', sheet_name='Customers')
regions = pd.read_excel('Sales.xlsx', sheet_name='Regions')
products = pd.read_excel('Sales.xlsx', sheet_name='Products')

# %%
# Renommer les colonnes pour éviter les conflits
customers_renamed = customers.rename(columns={
    'Customer Names': 'Customer Names_cust',
    'Size': 'Size_cust',
    'Capital': 'Capital_cust'
})

# Fusionner Sales Orders avec Customers
sales_orders = sales_orders.merge(
    customers_renamed,
    left_on='Customer Name Index',
    right_on='Customer Index',
    how='left'
)
sales_orders.drop(columns=['Customer Index'], inplace=True)

# Renommer les colonnes pour éviter les conflits
regions_renamed = regions.rename(columns={
    'Suburb': 'Suburb_reg',
    'City': 'City_reg',
    'postcode': 'postcode_reg',
    'Longitude': 'Longitude_reg',
    'Latitude': 'Latitude_reg',
    'Full Address': 'Full Address_reg'
})

# Fusionner Sales Orders avec Regions
sales_orders = sales_orders.merge(
    regions_renamed,
    left_on='Delivery Region Index',
    right_on='Index',
    how='left'
)
sales_orders.drop(columns=['Index'], inplace=True)

# Renommer les colonnes pour éviter les conflits
products_renamed = products.rename(columns={
    'Product Name': 'Product Name_prod'
})

# Fusionner Sales Orders avec Products
sales_orders = sales_orders.merge(
    products_renamed,
    left_on='Product Description Index',
    right_on='Index',
    how='left'
)
sales_orders.drop(columns=['Index'], inplace=True)

# %%
# Créer un DataFrame de dates couvrant la période des données
start_date = sales_orders['OrderDate'].min()
end_date = sales_orders['OrderDate'].max()
date_range = pd.date_range(start=start_date, end=end_date)

date_table = pd.DataFrame(date_range, columns=['Date'])
date_table['Year'] = date_table['Date'].dt.year
date_table['Month'] = date_table['Date'].dt.month
date_table['Quarter'] = date_table['Date'].dt.to_period('Q').astype(str)
date_table['DayOfWeek'] = date_table['Date'].dt.day_name()

# %%
# Calcul des ventes totales
sales_orders['Sales'] = sales_orders['Order Quantity'] * sales_orders['Unit Selling Price']
total_sales = sales_orders['Sales'].sum()

# Calcul des bénéfices totaux
sales_orders['Profit'] = (sales_orders['Unit Selling Price'] - sales_orders['Unit Cost']) * sales_orders['Order Quantity']
total_profit = sales_orders['Profit'].sum()

# Calcul de la marge bénéficiaire
profit_margin = total_profit / total_sales if total_sales != 0 else 0

# Comparaison des ventes de cette année avec l'année précédente
current_year = sales_orders['OrderDate'].dt.year.max()
previous_year = current_year - 1

sales_current_year = sales_orders[sales_orders['OrderDate'].dt.year == current_year].groupby('Product Name_prod')['Sales'].sum()
sales_last_year = sales_orders[sales_orders['OrderDate'].dt.year == previous_year].groupby('Product Name_prod')['Sales'].sum()
sales_comparison = pd.DataFrame({
    'Current Year': sales_current_year,
    'Last Year': sales_last_year
}).fillna(0)
sales_comparison['Difference'] = sales_comparison['Current Year'] - sales_comparison['Last Year']
sales_comparison['Change (%)'] = (sales_comparison['Difference'] / sales_comparison['Last Year']) * 100

# %%
plt.figure(figsize=(12, 6))
sns.barplot(data=sales_orders, x='Product Name_prod', y='Sales')
plt.title('Ventes par produit')
plt.xticks(rotation=90)
plt.show()

# %%
monthly_sales = sales_orders.groupby([sales_orders['OrderDate'].dt.year, sales_orders['OrderDate'].dt.month])['Sales'].sum().unstack()
monthly_sales.plot(kind='bar', figsize=(12, 6))
plt.title('Ventes par mois')
plt.xlabel('Année')
plt.ylabel('Ventes')
plt.show()

# %%
top_cities = sales_orders.groupby('City_reg')['Sales'].sum().nlargest(5)
top_cities.plot(kind='bar', figsize=(12, 6))
plt.title('Top 5 des villes par ventes')
plt.show()

# %%
channel_profit = sales_orders.groupby('Channel')['Profit'].sum()
channel_profit.plot(kind='bar', figsize=(12, 6))
plt.title('Bénéfices par canal')
plt.show()

# %%
top_clients = sales_orders.groupby('Customer Names_cust')['Sales'].sum().nlargest(5)
top_clients.plot(kind='bar', figsize=(12, 6))
plt.title('Top 5 des ventes par client')
plt.show()

# %%
last_clients = sales_orders.groupby('Customer Names_cust')['Sales'].sum().nsmallest(5)
last_clients.plot(kind='bar', figsize=(12, 6))
plt.title('Derniers 5 clients par ventes')
plt.show()

# %%
# Créer une carte centrée sur la moyenne des coordonnées
map_center = [sales_orders['Latitude_reg'].mean(), sales_orders['Longitude_reg'].mean()]
sales_map = folium.Map(location=map_center, zoom_start=10)

# Ajouter des marqueurs pour chaque vente
for _, row in sales_orders.iterrows():
    folium.Marker(
        location=[row['Latitude_reg'], row['Longitude_reg']],
        popup=f"{row['City_reg']}: {row['Sales']}",
        icon=folium.Icon(color='blue')
    ).add_to(sales_map)

sales_map.save('sales_map.html')

# %%
# Créer une carte centrée sur la moyenne des coordonnées
map_center = [sales_orders['Latitude_reg'].mean(), sales_orders['Longitude_reg'].mean()]
sales_map = folium.Map(location=map_center, zoom_start=10)

# Ajouter des marqueurs pour chaque vente
for _, row in sales_orders.iterrows():
    folium.Marker(
        location=[row['Latitude_reg'], row['Longitude_reg']],
        popup=f"{row['City_reg']}: {row['Sales']}",
        icon=folium.Icon(color='blue')
    ).add_to(sales_map)

sales_map.save('sales_map.html')
