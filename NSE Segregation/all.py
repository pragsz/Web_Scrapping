

# ## general,ytd,current, balance sheet, segment wworking fine
# import os
# import pandas as pd

# # Define paths
# folder_path = r"D:\NSE\excelfile"
# output_folder = r"D:\NSE\output"
# os.makedirs(output_folder, exist_ok=True)

# # List of all files in the folder, excluding temporary files
# files = [f for f in os.listdir(folder_path) if (f.endswith('.xlsx') or f.endswith('.xls')) and not f.startswith('~$')]

# # Define the extraction function for general data
# def Extract_General_Data(df):
#     start_value = 'ScripCode'
#     end_value = 'NatureOfReportStandaloneConsolidated'

#     start_index = df[df['Element Name'] == start_value].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         end_index = df[df['Element Name'] == end_value].index
#         if not end_index.empty:
#             end_index = end_index[0]
#             general_data = df.loc[start_index:end_index]
#             general_data = general_data[['Element Name', 'Unit', 'Fact Value']]
#             return general_data
#     return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

# # Modify the existing extraction functions
# def Extract_Income_Statement_Current(df):
#     general_data = Extract_General_Data(df)
    
#     start_value_col1 = 'RevenueFromOperations'
#     start_value_col2 = 'OneD'
#     end_value_col1 = 'DilutedEarningsLossPerShareFromContinuingAndDiscontinuedOperations'
#     end_value_col2 = 'OneD'
    
#     start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
#         if not end_index.empty:
#             end_index = end_index[0]
#             specific_data = df.loc[start_index:end_index]
#             specific_data = specific_data[['Element Name', 'Unit', 'Fact Value']]
#             # Combine general data and specific data
#             final_data = pd.concat([general_data, specific_data], ignore_index=True)
#             return final_data
#     return general_data  # Return only general data if specific data is not found

# def Extract_Income_Statement_YTD(df):
#     general_data = Extract_General_Data(df)
    
#     start_value_col1 = 'RevenueFromOperations'
#     start_value_col2 = 'FourD'
#     end_value_col1 = 'DilutedEarningsLossPerShareFromContinuingAndDiscontinuedOperations'
#     end_value_col2 = 'FourD'
    
#     start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
#         if not end_index.empty:
#             end_index = end_index[0]
#             specific_data = df.loc[start_index:end_index]
#             specific_data = specific_data[['Element Name', 'Unit', 'Fact Value']]
#             # Combine general data and specific data
#             final_data = pd.concat([general_data, specific_data], ignore_index=True)
#             return final_data
#     return general_data  # Return only general data if specific data is not found

# def Extract_Balance_Sheet(df):
#     general_data = Extract_General_Data(df)
    
#     start_value_col1 = 'PropertyPlantAndEquipment'
#     start_value_col2 = 'OneI'
#     end_value_col1 = 'EquityAndLiabilities'
#     end_value_col2 = 'OneI'
    
#     start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
#         if not end_index.empty:
#             end_index = end_index[0]
#             specific_data = df.loc[start_index:end_index]
#             specific_data = specific_data[['Element Name', 'Unit', 'Fact Value']]
#             # Combine general data and specific data
#             final_data = pd.concat([general_data, specific_data], ignore_index=True)
#             return final_data
#     return general_data  # Return only general data if specific data is not found


# def Extract_Segment(df):
#     # Extract general data
#     general_data = Extract_General_Data(df)
    
#     start_value_col1 = 'DescriptionOfReportableSegment'
#     end_value_col1 = 'NetSegmentLiabilities'
    
#     # Find the indices for the start and end of the segment
#     start_index = df[df['Element Name'] == start_value_col1].index
#     end_index = df[df['Element Name'] == end_value_col1].index
    
#     if not start_index.empty:
#         start_index = start_index[0]
#         if not end_index.empty:
#             end_index = end_index[0]
#             # Extract the rows between the start and end indices
#             edited_df = df.loc[start_index:end_index]
#             # Filter out rows where the 'Unit' column is 'OneD'
#             edited_df = edited_df[edited_df['Unit'] != 'OneD']
#             # Select only the required columns
#             edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
#             # Combine the general data with the specific segment data
#             final_data = pd.concat([general_data, edited_df], ignore_index=True)
#             return final_data

#     return general_data  # Return only general data if specific data is not found




# # Function to save results
# def save_results():
#     function_dict = {
#         'Extract_Income_Statement_Current': Extract_Income_Statement_Current,
#         'Extract_Income_Statement_YTD': Extract_Income_Statement_YTD,
#         'Extract_Balance_Sheet': Extract_Balance_Sheet,
#         'Extract_Segment': Extract_Segment,
#     }
    
#     for func_name, func in function_dict.items():
#         file_path = os.path.join(output_folder, f"{func_name}.xlsx")
#         with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
#             sheet_created = False  # Track if any sheets are created
#             for file in files:
#                 company_name = os.path.splitext(file)[0]
#                 df = pd.read_excel(os.path.join(folder_path, file))
#                 result_df = func(df)
#                 if not result_df.empty:
#                     result_df.to_excel(writer, sheet_name=company_name[:31], index=False)
#                     sheet_created = True  # Set to True if at least one sheet is created
#                     print(f"Data for {company_name} saved in {func_name}.xlsx")
#             if not sheet_created:
#                 # If no sheets were created, don't save an empty workbook
#                 print(f"No data found for {func_name}, skipping file creation.")
#                 writer.book = None  # Prevents saving the empty workbook

# # Run the function to save results
# save_results()

import os
import pandas as pd

# Define paths
folder_path = r"D:\NSE\excelfile"
output_folder = r"D:\NSE\output"
os.makedirs(output_folder, exist_ok=True)

# List of all files in the folder
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

# Constants
GENERAL_DATA_START = 'ScripCode'
GENERAL_DATA_END = 'NatureOfReportStandaloneConsolidated'

# Utility functions
def extract_dataframe_segment(df, start_value, end_value, unit_value=None):
    start_index = df[df['Element Name'] == start_value].index
    if not start_index.empty:
        start_index = start_index[0]
        end_index = df[df['Element Name'] == end_value].index
        if not end_index.empty:
            end_index = end_index[0]
            segment_df = df.loc[start_index:end_index]
            if unit_value:
                segment_df = segment_df[segment_df['Unit'] == unit_value]
            return segment_df[['Element Name', 'Unit', 'Fact Value']]
    return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

def extract_general_data(df):
    return extract_dataframe_segment(df, GENERAL_DATA_START, GENERAL_DATA_END)

# Define the extraction functions
def extract_income_statement(df, unit):
    general_data = extract_general_data(df)
    income_statement = extract_dataframe_segment(
        df, 'RevenueFromOperations', 
        'DilutedEarningsLossPerShareFromContinuingAndDiscontinuedOperations', 
        unit
    )
    return pd.concat([general_data, income_statement], ignore_index=True)

def extract_balance_sheet(df):
    general_data = extract_general_data(df)
    balance_sheet = extract_dataframe_segment(
        df, 'PropertyPlantAndEquipment', 
        'EquityAndLiabilities', 
        'OneI'
    )
    return pd.concat([general_data, balance_sheet], ignore_index=True)

def extract_equity_adjustment(df):
    return extract_dataframe_segment(df, 'DescriptionOfItemThatWillNotBeReclassifiedToProfitAndLoss', None)

def extract_cfs(df):
    general_data = extract_general_data(df)
    cfs_data = extract_dataframe_segment(
        df, 'AdjustmentsForFinanceCosts', 
        'IncreaseDecreaseInCashAndCashEquivalents', 
        'FourD'
    )
    return pd.concat([general_data, cfs_data], ignore_index=True)

def extract_segment(df):
    segment_df = extract_dataframe_segment(df, 'DescriptionOfReportableSegment', 'NetSegmentLiabilities')
    income_statement = extract_income_statement(segment_df[segment_df['Unit'] == 'OneD'], 'OneD')
    return pd.concat([segment_df, income_statement], ignore_index=True)

def extract_adjustment2(df):
    def extract_value(element_name, unit_name):
        filtered_df = df[(df['Element Name'] == element_name) & (df['Unit'] == unit_name)]
        if not filtered_df.empty:
            value = filtered_df['Fact Value'].iloc[0]
            try:
                return float(str(value).replace(',', ''))
            except ValueError:
                return 0
        return 0

    values = [
        extract_value('AmountOfItemThatWillBeReclassifiedToProfitAndLoss', 'OneD'),
        extract_value('IncomeTaxRelatingToItmesThatWillBeReclassifiedToProfitOrLoss', 'OneD'),
        extract_value('AmountOfItemThatWillNotBeReclassifiedToProfitAndLoss', 'OneD'),
        extract_value('IncomeTaxRelatingToItmesThatWillNotBeReclassifiedToProfitAndLoss', 'OneD')
    ]
    
    net_total = values[0] + values[1] - values[2] - values[3]
    
    data = {
        'Element Name': [
            'Amount of Item That Will Be Reclassified To Profit Or Loss',
            'Amount of Item That Will Not Be Reclassified To Profit and Loss',
            'Income Tax Relating To Items That Will Not Be Reclassified To Profit Or Loss',
            'Net Total'
        ],
        'Fact Value': values + [net_total]
    }
    
    return pd.DataFrame(data)

# Dictionary of extraction functions
extraction_functions = {
    'Income_Statement_Current': lambda df: extract_income_statement(df, 'OneD'),
    'Income_Statement_YTD': lambda df: extract_income_statement(df, 'FourD'),
    'Balance_Sheet': extract_balance_sheet,
    'Equity_Adjustment': extract_equity_adjustment,
    'CFS': extract_cfs,
    'Segment': extract_segment,
    'Adjustment2': extract_adjustment2
}

# Save each function's results into a single sheet per company
def save_results():
    for func_name, func in extraction_functions.items():
        file_path = os.path.join(output_folder, f"{func_name}.xlsx")
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            sheet_created = False  # Track if any sheet is created
            for file in files:
                company_name = os.path.splitext(file)[0]
                df = pd.read_excel(os.path.join(folder_path, file))
                result_df = func(df)
                
                if not result_df.empty:
                    sheet_name = company_name[:31]  # Sheet names are limited to 31 characters in Excel
                    result_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    sheet_created = True
                    print(f"Data for {company_name} saved in {func_name}.xlsx")
            
            if not sheet_created:
                # If no sheets were created, add a dummy sheet
                pd.DataFrame({"Message": ["No data available"]}).to_excel(writer, sheet_name="NoData", index=False)

save_results()


# Combine company sheets
def combine_company_sheets(input_dir="excelfile", workbook_dir="output", output_dir="combined"):
    os.makedirs(output_dir, exist_ok=True)
    dir_list = os.listdir(input_dir)
    formatted_list = [x[:31] for x in dir_list]
    company_list = list(dict.fromkeys([x.split('_')[-1][:-5] for x in dir_list]))
    file_list = os.listdir(workbook_dir)
    
    for workbook in file_list:
        xls = pd.ExcelFile(os.path.join(workbook_dir, workbook))
        print(f"Processing workbook: {workbook}")
        company_df = []
        
        for company in company_list:
            company_sheets = []
            for i, name in enumerate(formatted_list):
                if dir_list[i].find(company) > 0:
                    try:
                        df = pd.read_excel(xls, name)[["Element Name", "Fact Value"]]
                        df.rename(columns={"Fact Value": name.split('_')[0]}, inplace=True)
                        company_sheets.append(df)
                    except Exception as e:
                        print(f"Error reading sheet: {e}")
            
            if company_sheets:
                base_df = max(company_sheets, key=lambda df: len(df['Element Name'].unique()))[['Element Name']].drop_duplicates().reset_index(drop=True)
                for df in company_sheets:
                    base_df = pd.merge(base_df, df, on='Element Name', how='left')
                company_df.append(base_df)
        
        output_file = os.path.join(output_dir, workbook)
        try:
            with pd.ExcelWriter(output_file) as writer:
                for i, df in enumerate(company_df):
                    df.to_excel(writer, sheet_name=company_list[i][:31], index=False)
            print(f"Workbook saved: {output_file}")
        except Exception as e:
            print(f"Error writing to Excel: {e}")

combine_company_sheets()


