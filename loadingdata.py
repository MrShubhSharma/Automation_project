import pandas as pd

# Load the Excel files
file1_path = 'TL_filteration_file.xlsx'
file2_path = 'GUJ- LINK LIST OCT New.xlsx'

df_tl = pd.read_excel(file1_path)
df_guj = pd.read_excel(file2_path)

# Specify the columns to fill if missing for A END NODE
columns_to_fill_a = ['A END NSSID', 'A END ZONE', 'LATLONG A', 'A END GENERIC NAME-NEW']

# Specify the columns to fill if missing for B END NODE
columns_to_fill_b = ['B END NSSID', 'B END ZONE', 'LATLONG B', 'B END GENERIC NAME-NEW']

# Ensure 'A END NODE' and 'B END NODE' do not have NaN values
df_guj = df_guj[df_guj['A END NODE'].notna() & df_guj['B END NODE'].notna()]

# Function to fill missing values
def fill_missing_values(df, node_col, columns_to_fill, end_nodes):
    for end_node in end_nodes:
        matches = df[df[node_col].str.startswith(end_node, na=False)]

        if not matches.empty:
            complete_entry = matches.dropna(subset=columns_to_fill).iloc[0] if not matches.dropna(subset=columns_to_fill).empty else None

            if complete_entry is not None:
                for col in columns_to_fill:
                    if pd.isna(complete_entry[col]):
                        continue
                    df.loc[(df[node_col].str.startswith(end_node) & df[col].isna()), col] = complete_entry[col]

# Fill missing values for A END NODE
fill_missing_values(df_guj, 'A END NODE', columns_to_fill_a, df_tl['NODE NAME'].unique())

# Fill missing values for B END NODE
fill_missing_values(df_guj, 'B END NODE', columns_to_fill_b, df_tl['NODE NAME'].unique())

# Function to fill corresponding values
def fill_corresponding_values(df):
    for index, row in df.iterrows():
        a_end_node = row['A END NODE']
        b_end_node = row['B END NODE']

        # If A END NODE has values and B END NODE is the same but has missing values
        if a_end_node in df['B END NODE'].values:
            for col in columns_to_fill_b:
                if pd.isna(row[col]) and not pd.isna(row[columns_to_fill_a[columns_to_fill_b.index(col)]]):
                    df.at[index, col] = row[columns_to_fill_a[columns_to_fill_b.index(col)]]

        # If B END NODE has values and A END NODE is the same but has missing values
        if b_end_node in df['A END NODE'].values:
            for col in columns_to_fill_a:
                if pd.isna(row[col]) and not pd.isna(row[columns_to_fill_b[columns_to_fill_a.index(col)]]):
                    df.at[index, col] = row[columns_to_fill_b[columns_to_fill_a.index(col)]]

# Fill missing values based on corresponding node
fill_corresponding_values(df_guj)

# Identify rows with missing values in B END columns and corresponding A END columns
mask_missing_b = df_guj[columns_to_fill_b].isna().any(axis=1)
mask_missing_a = df_guj[columns_to_fill_a].isna().any(axis=1)

# Create a new DataFrame for rows that meet the criteria
rows_with_issues = df_guj[mask_missing_b | mask_missing_a]

# Save the rows with missing data to a new Excel file
if not rows_with_issues.empty:
    output_issues_path = 'Rows_With_Missing_Data.xlsx'
    rows_with_issues.to_excel(output_issues_path, index=False)
    print(f"Rows with missing data saved to {output_issues_path}")
else:
    print("No rows with missing data found.")

# Save the modified DataFrame to a new Excel file
output_path = 'Updated_GUJ_Link_List.xlsx'
df_guj.to_excel(output_path, index=False)

# Create a new DataFrame with only rows where all specified columns have values
complete_data_df = df_guj.dropna(subset=['A END NSSID', 'A END ZONE', 'LATLONG A', 
                                           'A END GENERIC NAME-NEW', 'B END NSSID', 
                                           'B END ZONE', 'LATLONG B', 'B END GENERIC NAME-NEW'])

# Save this DataFrame to a new Excel file
complete_data_output_path = 'Complete_Columns_Data.xlsx'
complete_data_df.to_excel(complete_data_output_path, index=False)
print(f"Complete columns data saved to {complete_data_output_path}")
