# @title Default title text
# Step 1: Install required packages (if needed)
!pip install pandas openpyxl

# Step 2: Import libraries
import pandas as pd
from google.colab import files

# Step 3: Upload your CSV file
uploaded = files.upload()

# Step 4: Process the uploaded file
for file_name in uploaded.keys():
    # Load the CSV file, skipping the first 13 rows
    df = pd.read_csv(file_name, skiprows=13)

    # Clean column names
    df.columns = df.columns.str.replace('\u00A0', ' ').str.strip()

    # Identify the correct column names
    time_col = df.columns[0]
    usage_col = df.columns[1]

    # Extract start time from 'Energy Delivered time period'
    df['Time Period'] = df[time_col].str.split(' to ').str[0]

    # Convert 'Usage Delivered' to numeric
    df['Usage Delivered'] = pd.to_numeric(df[usage_col], errors='coerce')

    # Create a dictionary to store first and second usage values
    usage_dict = {}
    for _, row in df.iterrows():
        time_period = row['Time Period']
        usage = row['Usage Delivered']
        if pd.notna(usage):
            if time_period not in usage_dict:
                usage_dict[time_period] = [usage, None]
            elif usage_dict[time_period][1] is None:
                usage_dict[time_period][1] = usage

    # Calculate the difference
    result = []
    for time_period, usages in usage_dict.items():
        if usages[1] is not None:
            usage_diff = usages[0] - usages[1]
            result.append([time_period, usage_diff])

    # Create result DataFrame
    result_df = pd.DataFrame(result, columns=['Time Period', 'Usage Difference'])

    # Save to Excel
    output_file = 'Modified_SCE_Usage.xlsx'
    result_df.to_excel(output_file, index=False)

    # Download the file
    files.download(output_file)
