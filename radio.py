df_loaded.columns = df_loaded.columns.str.replace(' ', '_')
df = df_loaded.copy()
df['Shipment_n'] = df['Shipment_n'].astype(str)
df['Gate_Out_D'] = pd.to_datetime(df['Gate_Out_D'])
df['City_Code'] = df['City_Code'].astype(str)
df['Gate_Out_D'] = pd.to_datetime(df['Gate_Out_D'])
filter_year = 2021
# Filter the DataFrame based on the year in column 'C'Gate Out D
filtered_df = df[df['Gate_Out_D'].dt.year == filter_year]
filtered_df = filtered_df[filtered_df['Truck_No'].astype(str).str.startswith(('U','D','F'))]
filtered_df.reset_index(drop=True, inplace=True)
filtered_df.index += 1
filtered_df['Shipment_n'] = filtered_df['Shipment_n'].astype(str)
# filtered_df['Gate_Out_D'] = filtered_df.to_datetime(df['Gate_Out_D'])
filtered_df['City_Code'] = filtered_df['City_Code'].astype(str)
filtered_df = filtered_df.sort_values(by='Gate_Out_D', ascending=True)
filtered_df.head()
filtered_df.reset_index(drop=True, inplace=True)
filtered_df.index += 1
df = filtered_df.copy()
df['Shipment_n'] = df['Shipment_n'].astype(str)
df['Gate_Out_D'] = pd.to_datetime(df['Gate_Out_D'])
df['City_Code'] = df['City_Code'].astype(str)
idx = df.groupby(['Shipment_n', 'Gate_Out_D'])['Base_Freig'].idxmax()
# Create a new DataFrame to map the highest Base Freig City Code to each combination
base_area_map = df.loc[idx, ['Shipment_n', 'Gate_Out_D', 'City_Code']].rename(columns={'City_Code': 'Base_Area'})
df['Shipment_n'] = df['Shipment_n'].astype(str)
df['Gate_Out_D'] = pd.to_datetime(df['Gate_Out_D'])
df['City_Code'] = df['City_Code'].astype(str)
# Merge this mapping back to the original DataFrame
df = df.merge(base_area_map, on=['Shipment_n', 'Gate_Out_D'], how='left')
df['Shipment_n'] = df['Shipment_n'].astype(str)
df['Gate_Out_D'] = pd.to_datetime(df['Gate_Out_D'])
df['City_Code'] = df['City_Code'].astype(str)
len(df)
# Group by 'Base City' and 'City', then aggregate counts
city_counts = df.groupby(['Base_Area', 'City_Code']).size().reset_index(name='Count')

# Sort by 'Base City' and descending 'Count' within each group
city_counts_sorted = city_counts.sort_values(by=['Base_Area', 'Count'], ascending=[True, False])

# Initialize a dictionary to store top 3 cities for each base city
top_cities_by_base_city = {}

# Iterate through each base city group
for base_city, group in city_counts_sorted.groupby('Base_Area'):
    # Get the top 3 cities for the current base city
    top_cities = group[['City_Code', 'Count']].to_records(index=False).tolist()
    top_cities_by_base_city[base_city] = top_cities

# Print or use the top 3 cities for each base city
for base_city, top_cities in top_cities_by_base_city.items():
    print(f"Base_Area: {base_city}")
    for rank, (city, count) in enumerate(top_cities, start=1):
        print(f"  {rank}. {city} - Count: {count}")
    print()
# print(top_cities_by_base_city)

# Group by 'Base_Area' and 'City Code', then aggregate counts
city_counts = df.groupby(['Base_Area', 'City_Code']).size().reset_index(name='Count')

# Filter out cities with counts less than 5
# city_counts = city_counts[city_counts['Count'] >= 10]

# Sort by 'Base_Area' and descending 'Count' within each group
city_counts_sorted = city_counts.sort_values(by=['Base_Area', 'Count'], ascending=[True, False])

# Initialize a dictionary to store top cities for each Base_Area
top_cities_by_base_area = {}

# Iterate through each Base_Area group
for base_area, group in city_counts_sorted.groupby('Base_Area'):
    # Get the top cities for the current Base_Area and convert to list
    top_cities = group['City_Code'].tolist()
    top_cities_by_base_area[base_area] = top_cities

# Print the result
print(top_cities_by_base_area)
all_values = []
for value_list in top_cities_by_base_city.values():
    all_values.extend(value_list)

# Convert the combined list to a set to get unique values
# unique_values = set(all_values)

# # Get the count of unique values
# unique_count = len(unique_values)

# # Print the unique count and unique values
# print("Count of unique values:", unique_count)
# print("Unique values:", unique_values)
