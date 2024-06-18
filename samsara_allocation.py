import os
import pandas as pd
import numpy as np
from pandas.tseries.offsets import BDay

def read_and_combine_excel_files(input_folder_path):
    # Read and combine all Excel files in the folder, adding a filename column
    dfs = []
    for file in os.listdir(input_folder_path):
        if file.endswith(".xlsx"):
            file_path = os.path.join(input_folder_path, file)
            df = pd.read_excel(file_path)
            df['Filename'] = os.path.splitext(file)[0]
            dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def clean_data(df):
    # Fill missing drivers, format dates, and add job details
    df['Driver'] = df['Driver'].fillna('Unassigned')
    df['Arrival'] = pd.to_datetime(df['Arrival']).dt.strftime('%m/%d/%Y')
    df['Departure'] = pd.to_datetime(df['Departure']).dt.strftime('%m/%d/%Y')
    df['Hours'] = df['Time on Site (Minutes)'] / 60
    df[['Job', 'Job Name']] = df['Filename'].apply(lambda x: split_filename(x)).tolist()
    df['Job Name'] = df['Job Name'].str.replace(r'\.xlsx$', '', regex=True)
    df.drop(['Filename', 'GPS Distance Traveled (mi)', ], axis=1, inplace=True)
    return df

def split_filename(filename):
    # Split filename into job and job name
    parts = filename.split('-', 1)
    job = parts[0].strip() if parts[0].strip().isdigit() else "Overhead"
    job_name = parts[1].strip() if len(parts) > 1 else filename
    return job, job_name

def aggregate_data(df):
    # Group by key columns and aggregate data
    aggregated_df = df.groupby(['Asset', 'Driver', 'Job', 'Job Name', 'Arrival', 'Departure']).agg({
        'Time on Site (Minutes)': 'sum',
        'Hours': 'sum'
    }).reset_index()
    aggregated_df['Days'] = aggregated_df['Hours'].apply(lambda x: max(1, int(x / 8)) if x > 4 else 0)
    return aggregated_df

def calculate_business_days(row):
    if row['Arrival'] == row['Departure']:
        return 1 if row['Hours'] > 4 else 0
    else:
        return len(pd.date_range(start=row['Arrival'], end=row['Departure'], freq=BDay()))

def calculate_weekend_days(row):
    return len(pd.date_range(start=row['Arrival'], end=row['Departure'], freq='W-SAT')) + len(pd.date_range(start=row['Arrival'], end=row['Departure'], freq='W-SUN'))

def analysis_data(aggregated_df, cleaned_df):
    # Calculate Job Total Hours rounded to 2 digits
    job_hours = cleaned_df.groupby(['Asset', 'Job']).agg({'Hours': 'sum'}).reset_index()
    job_hours['Hours'] = job_hours['Hours'].round(2)
    aggregated_df = aggregated_df.merge(job_hours, on=['Asset', 'Job'], how='left', suffixes=('', ' Job Total'))

    # Calculate Total Jobs for each asset
    total_jobs = cleaned_df[cleaned_df['Job'] != 'Overhead'].groupby('Asset')['Job'].nunique().reset_index()
    total_jobs.columns = ['Asset', 'Total Jobs']
    aggregated_df = aggregated_df.merge(total_jobs, on='Asset', how='left')

    # Convert Arrival and Departure back to datetime for calculation
    aggregated_df['Arrival'] = pd.to_datetime(aggregated_df['Arrival'], format='%m/%d/%Y')
    aggregated_df['Departure'] = pd.to_datetime(aggregated_df['Departure'], format='%m/%d/%Y')

    # Calculate Business Days and Weekend Days
    aggregated_df['Business Days'] = aggregated_df.apply(calculate_business_days, axis=1)
    aggregated_df['Weekend Days'] = aggregated_df.apply(calculate_weekend_days, axis=1)

    return aggregated_df

def summarize_analysis_data(analysis_df):
    # Group and aggregate data as specified
    summary_df = analysis_df.groupby(['Asset', 'Driver', 'Job', 'Job Name', 'Hours Job Total', 'Total Jobs']).agg({
        'Hours': lambda x: round(x.sum(), 2),
        'Business Days': 'sum',
        'Weekend Days': 'sum',
    }).reset_index()

    # Add Allocation Pct based on given logic
    def calculate_allocation_pct(row):
        if row['Total Jobs'] >= 4:
            return 0  # Allocation is "Overhead" thus 0%
        elif row['Business Days'] >= 14:
            return 1.00  # Allocation is 100%
        elif row['Job'] == 'Overhead':
            return 0  # Allocation Pct for Overhead is 0%
        else:
            return round(row['Business Days'] / 21, 2)  # Remaining cases

    # Apply function to calculate Allocation Pct
    summary_df['Allocation Pct'] = summary_df.apply(calculate_allocation_pct, axis=1)

    # Enforce no asset's total allocation exceeds 1
    max_allocation = summary_df.groupby('Asset')['Allocation Pct'].transform('max')
    summary_df['Allocation Pct'] = summary_df.apply(lambda x: x['Allocation Pct'] if max_allocation[x.name] == x['Allocation Pct'] else 0, axis=1)

    # Calculate Pickup Repair
    summary_df['Pickup Repair'] = summary_df['Allocation Pct'].apply(lambda x: round(x * 620, 2))

    return summary_df

def prepare_final_tab(summary_df, vehicle_lookup_path):
    summary_df['Asset'] = summary_df['Asset'].astype(str).str.strip()  # Ensure Asset is a string and strip whitespace
    vehicle_names = read_vehicle_names(vehicle_lookup_path)

    # Debug prints to check the data before merging
    print("Summary DataFrame Before Merge:", summary_df[['Asset']].head())
    print("Vehicle Names DataFrame:", vehicle_names.head())

    # Perform the merge
    final_df = pd.merge(summary_df, vehicle_names, on='Asset', how='left')

    # Debug print to check the merge result
    print("Final DataFrame After Merge:", final_df[['Asset', 'Vehicle Name']].head())

    # Add blank columns and rearrange as necessary
    final_df['Vehicle'] = ''
    final_df['Fuel Cost'] = ''
    final_df['Total'] = ''
    final_df['Bill Rate'] = ''

    final_df = final_df[[
        'Asset', 'Vehicle Name', 'Driver', 'Business Days', 'Vehicle', 
        'Fuel Cost', 'Pickup Repair', 'Total', 'Bill Rate', 'Hours', 
        'Allocation Pct', 'Job', 'Job Name', 'Total Jobs'
    ]]

    return final_df

def read_vehicle_names(file_path):
    vehicle_df = pd.read_excel(file_path, usecols=[0, 1], header=None, names=['Asset', 'Vehicle Name'])
    vehicle_df['Asset'] = vehicle_df['Asset'].astype(str).str.strip()  # Convert to string and strip whitespace
    print("Vehicle Names Loaded:", vehicle_df.head())  # Debug print to check data
    return vehicle_df

def write_to_excel(df_dict, output_path):
    # Write multiple dataframes to different sheets in an Excel file
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)

def main(input_folder_path, output_folder_path, output_file_name, vehicle_lookup_path):
    combined_df = read_and_combine_excel_files(input_folder_path)
    cleaned_df = clean_data(combined_df)
    aggregated_df = aggregate_data(cleaned_df)
    analysis_df = analysis_data(aggregated_df, cleaned_df)
    summary_df = summarize_analysis_data(analysis_df)
    final_df = prepare_final_tab(summary_df, vehicle_lookup_path)

    output_path = os.path.join(output_folder_path, output_file_name)
    write_to_excel({
        'Original': combined_df, 
        'Clean_Data': cleaned_df, 
        'Aggregated_Data': aggregated_df,
        'Analysis': analysis_df,
        'Summary': summary_df,
        'Final': final_df  # Write the final tab
    }, output_path)

# Parameters
input_folder_path = r'C:\Users\gmakakaufaki\Sukut\Accounting - GL\01 General Ledger\01 SCLLC\Month-End Tasks\01 Vehicle Job Allocation\2024-05\Samsara Data'
output_folder_path = r'C:\Users\gmakakaufaki\Sukut\Accounting - GL\01 General Ledger\01 SCLLC\Month-End Tasks\01 Vehicle Job Allocation\2024-05'
output_file_name = 'Samsara_combined_excel.xlsx'
vehicle_lookup_path = r'C:\Users\gmakakaufaki\Sukut\Accounting - GL\01 General Ledger\01 SCLLC\Month-End Tasks\01 Vehicle Job Allocation\Supporting Files\01 EM Equipment List - Summary.xlsx'

# Run the main function
main(input_folder_path, output_folder_path, output_file_name, vehicle_lookup_path)

