# import os
# import pandas as pd

# # Get the list of all files and directories
# path = "excel"
# dir_list = os.listdir(path)
# print("Files and directories in '", path, "' :")
# # prints all files
# formatted_list = [x[:31] for x in dir_list]
# company_list = list(dict.fromkeys([x[(x.rfind("_"))+1:-5] for x in dir_list]))
# print(company_list)

# file_list = os.listdir("workbook")


# for files in file_list:
#     xls = pd.ExcelFile(f"workbook/{files}")
#     print(files)
#     company_df = []
#     for company in company_list:
        
#         company_sheets = []
#         for i in range(len(dir_list)):
#             if dir_list[i].find(company) > 0:
#                 try:
#                     name=formatted_list[i]
#                     print(name)
#                     df = (pd.read_excel(xls, name))[["Element Name", "Fact Value"]]
              
#                     df.rename(columns={"Fact Value": name[:name.rfind("_")]}, inplace=True)
#                     company_sheets.append(df)
#                     print(df.head())
#                 except Exception as e:
#                     print(f"Error reading sheet: {e}")
#         if len(company_sheets) > 0 :
#             final_df = company_sheets[0]
#             for j in range(1, len(company_sheets)):
#                 print(final_df.shape)
#                 final_df = pd.merge(final_df, company_sheets[j],  on="Element Name",  how='inner')

#             print(final_df.shape)
#             company_df.append(final_df)
#     try:
#         with pd.ExcelWriter(f"combined/{files}") as writer:
        
#             # use to_excel function and specify the sheet_name and index 
#             # to store the dataframe in specified sheet
#             for i in range(len(company_df)):
#                 company_df[i].to_excel(writer, sheet_name=company_list[i], index=False)
#     except Exception as e:
#         print(f"Error writing to excel: {e}") 


# code for merging the data uccefully 
# import os
# import pandas as pd

# # Define paths
# folder_path = r"D:\excel_workbook\excel"
# output_folder = r"D:\excel_workbook\workbook"
# os.makedirs(output_folder, exist_ok=True)

# # List of all files in the folder
# files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

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


# # Define the extraction functions
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
#             edited_df = df.loc[start_index:end_index]
#             edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
#             final_data = pd.concat([general_data, edited_df], ignore_index=True)
#             return final_data
        
#     return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

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
#             edited_df = df.loc[start_index:end_index]
#             edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
#             final_data = pd.concat([general_data, edited_df], ignore_index=True)
#             return final_data
#     return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

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
#             edited_df = df.loc[start_index:end_index]
#             edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
#             final_data = pd.concat([general_data, edited_df], ignore_index=True)
#             return final_data
#     return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

# def Extract_Equity_Adjustment(df):
#     start_value_col1 = 'DescriptionOfItemThatWillNotBeReclassifiedToProfitAndLoss'  
#     start_index = df[(df['Element Name'] == start_value_col1)].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         edited_df = df.loc[start_index:]
#         edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
#         return edited_df
#     return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

# def Extract_CFS(df):
#     general_data = Extract_General_Data(df)
#     start_value_col1 = 'AdjustmentsForFinanceCosts'
#     start_value_col2 = 'FourD'
#     end_value_col1 = 'IncreaseDecreaseInCashAndCashEquivalents'
#     end_value_col2 = 'FourD'
    
#     start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
#         if not end_index.empty:
#             end_index = end_index[0]
#             edited_df = df.loc[start_index:end_index]
#             edited_df = edited_df[['Element Name', 'Fact Value']]
#             final_data = pd.concat([general_data, edited_df], ignore_index=True)
#             return final_data
#     return pd.DataFrame(columns=['Element Name', 'Fact Value'])

# def Extract_Segment(df):
#     start_value_col1 = 'DescriptionOfReportableSegment'
#     end_value_col1 = 'NetSegmentLiabilities'
    
#     # Identify the segment of interest
#     start_index = df[(df['Element Name'] == start_value_col1)].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         end_index = df[(df['Element Name'] == end_value_col1)].index
#         if not end_index.empty:
#             end_index = end_index[0]
#             segment_df = df.loc[start_index:end_index]
            
#             # Filter the rows where 'Unit' is 'OneD'
#             oned_rows = segment_df[segment_df['Unit'] == 'OneD']
#             if not oned_rows.empty:
#                 # Send these rows to the Extract_Income_Statement_Current function
#                 income_statement_df = Extract_Income_Statement_Current(oned_rows)
                
#                 # Remove 'OneD' rows from segment_df
#                 segment_df = segment_df[segment_df['Unit'] != 'OneD']
            
#             # Return the remaining segment_df
#             segment_df = segment_df[['Element Name', 'Unit', 'Fact Value']]
#             return segment_df

#     return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])


# def Extract_Adjustment2(df):
#     set_11 = 'AmountOfItemThatWillBeReclassifiedToProfitAndLoss'
#     set_12 = 'OneD'
#     set_21 = 'IncomeTaxRelatingToItmesThatWillBeReclassifiedToProfitOrLoss'
#     set_22 = 'OneD'
#     set_31 = 'AmountOfItemThatWillNotBeReclassifiedToProfitAndLoss'
#     set_32 = 'OneD'
#     set_41 = 'IncomeTaxRelatingToItmesThatWillNotBeReclassifiedToProfitAndLoss'
#     set_42 = 'OneD'

#     def extract_value(df, element_name, unit_name):
#         filtered_df = df[(df['Element Name'] == element_name) & (df['Unit'] == unit_name)]
#         if not filtered_df.empty:
#             value = filtered_df['Fact Value'].iloc[0]
#             try:
#                 return float(str(value).replace(',', ''))
#             except ValueError:
#                 return 0
#         return 0

#     value_1 = extract_value(df, set_11, set_12)
#     value_2 = extract_value(df, set_21, set_22)
#     value_3 = extract_value(df, set_31, set_32)
#     value_4 = extract_value(df, set_41, set_42)

#     # Calculate the net total
#     Ans1 = value_1 + value_2 - value_3 - value_4

#     # Create lists for DataFrame construction
#     element_names = [
#         'Amount of Item That Will Be Reclassified To Profit Or Loss',
#         'Amount of Item That Will Not Be Reclassified To Profit and Loss',
#         'Income Tax Relating To Items That Will Not Be Reclassified To Profit Or Loss',
#         'Net Total'
#     ]

#     fact_values = [value_1, value_2, value_3, value_4, Ans1]

#     # Ensure both lists have the same length
#     if len(element_names) != len(fact_values):
#         # If lengths don't match, adjust the lists to have the same length
#         min_length = min(len(element_names), len(fact_values))
#         element_names = element_names[:min_length]
#         fact_values = fact_values[:min_length]

#     # Create DataFrame
#     data = {
#         'Element Name': element_names,
#         'Fact Value': fact_values
#     }

#     edited_df = pd.DataFrame(data)
#     return edited_df

# files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]


# # Save each function's results into a single sheet per company
# def save_results():
#     function_dict = {
#         'Extract_Income_Statement_Current': Extract_Income_Statement_Current,
#         'Extract_Income_Statement_YTD': Extract_Income_Statement_YTD,
#         'Extract_Balance_Sheet': Extract_Balance_Sheet,
#         'Extract_Equity_Adjustment': Extract_Equity_Adjustment,
#         'Extract_CFS': Extract_CFS,
#         'Extract_Segment': Extract_Segment,
#         'Extract_Adjustment2': Extract_Adjustment2
#     }

#     for func_name, func in function_dict.items():
#         file_path = os.path.join(output_folder, f"{func_name}.xlsx")
#         with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
#             sheet_data = {}  # Dictionary to hold data for each company
#             has_data = False  # Flag to track if any data is processed
            
#             for file in files:
#                 file_name = os.path.splitext(file)[0]
#                 company_name = file_name[11:]  # Adjust index if date format changes
#                 df = pd.read_excel(os.path.join(folder_path, file))
#                 result_df = func(df)
#                 if not result_df.empty:
#                     has_data = True
#                     if company_name not in sheet_data:
#                         sheet_data[company_name] = []
#                     sheet_data[company_name].append(result_df)
            
#             # If no data for any function, add a placeholder sheet
#             if not has_data:
#                 placeholder_df = pd.DataFrame({'Element Name': [], 'Fact Value': []})
#                 placeholder_df.to_excel(writer, sheet_name='Placeholder', index=False)

#             # Write all accumulated data to a single sheet for each company
#             for company_name, dfs in sheet_data.items():
#                 combined_df = pd.concat(dfs, ignore_index=True)
#                 sheet_name = company_name[:31]
#                 combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
#                 print(f"Data for {company_name} saved in {func_name}.xlsx")

# # Run the function to save results
# save_results()


# def combine_company_sheets(input_dir="excel", workbook_dir="workbook", output_dir="combined"):
#     # Ensure output directory exists
#     os.makedirs(output_dir, exist_ok=True)

#     # Get the list of all files and directories in the input directory
#     dir_list = os.listdir(input_dir)
    
#     # Adjust the extraction logic to handle new filenames
#     company_list = [os.path.splitext(file)[0][11:] for file in dir_list]  # Extract company names assuming date is at start
    
#     # Get the list of workbook files
#     file_list = os.listdir(workbook_dir)
    
#     # Loop through each workbook file
#     for file in file_list:
#         xls = pd.ExcelFile(os.path.join(workbook_dir, file))
#         print(f"Processing workbook: {file}")
#         company_df = []
        
#         # Loop through each company
#         for company in company_list:
#             company_sheets = []
            
#             # Loop through each directory/file to match the company name
#             for i in range(len(dir_list)):
#                 if company in dir_list[i]:
#                     try:
#                         name = dir_list[i][:31]  # Truncate sheet name to 31 characters
#                         print(f"Processing sheet: {name}")
#                         # Read the Excel sheet and filter columns
#                         df = pd.read_excel(xls, name)[["Element Name", "Fact Value"]]
#                         df.rename(columns={"Fact Value": name[:name.rfind("_")]}, inplace=True)
#                         company_sheets.append(df)
#                     except Exception as e:
#                         print(f"Error reading sheet: {e}")
            
#             # Combine all sheets for the company
#             if company_sheets:
#                 # Determine the sheet with the maximum number of unique Element Names
#                 max_elements_sheet = max(company_sheets, key=lambda df: len(df['Element Name'].unique()))
                
#                 # Create a base dataframe with all unique Element Names
#                 base_df = max_elements_sheet[['Element Name']].drop_duplicates().reset_index(drop=True)
                
#                 # Merge all sheets based on the Element Name
#                 for df in company_sheets:
#                     # Ensure that the sheet has the correct columns
#                     df = df[['Element Name', df.columns[1]]]
#                     df.rename(columns={df.columns[1]: df.columns[1]}, inplace=True)
#                     base_df = pd.merge(base_df, df, on='Element Name', how='left')
                
#                 # Append the combined dataframe for this company
#                 company_df.append(base_df)
        
#         # Write the combined data into a new Excel file
#         try:
#             output_file = os.path.join(output_dir, file)
#             with pd.ExcelWriter(output_file) as writer:
#                 for i, df in enumerate(company_df):
#                     sheet_name = company_list[i][:31]  # Truncate sheet name to 31 characters
#                     df.to_excel(writer, sheet_name=sheet_name, index=False)
#             print(f"Workbook saved: {output_file}")
#         except Exception as e:
#             print(f"Error writing to Excel: {e}")

# # Call the function
# combine_company_sheets()



# code for merging the data uccefully 
import os
import pandas as pd

# Define paths
folder_path = r"D:\excel_workbook\excel"
output_folder = r"D:\excel_workbook\workbook"
os.makedirs(output_folder, exist_ok=True)

# List of all files in the folder
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

def Extract_General_Data(df):
    start_value = 'ScripCode'
    end_value = 'NatureOfReportStandaloneConsolidated'
 
    start_index = df[df['Element Name'] == start_value].index
    if not start_index.empty:
        start_index = start_index[0]
        end_index = df[df['Element Name'] == end_value].index
        if not end_index.empty:
            end_index = end_index[0]
            general_data = df.loc[start_index:end_index]
            general_data = general_data[['Element Name', 'Unit', 'Fact Value']]
            print("General data extracted:", general_data)
            return general_data
    return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])


# Define the extraction functions
def Extract_Income_Statement_Current(df):
    general_data = Extract_General_Data(df)
    
    start_value_col1 = 'RevenueFromOperations'
    start_value_col2 = 'OneD'
    end_value_col1 = 'DilutedEarningsLossPerShareFromContinuingAndDiscontinuedOperations'
    end_value_col2 = 'OneD'
    
    start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
    if not start_index.empty:
        start_index = start_index[0]
        end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
        if not end_index.empty:
            end_index = end_index[0]
            edited_df = df.loc[start_index:end_index]
            edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
            final_data = pd.concat([general_data, edited_df], ignore_index=True)
            return final_data
        
    return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

def Extract_Income_Statement_YTD(df):
    general_data = Extract_General_Data(df)
    start_value_col1 = 'RevenueFromOperations'
    start_value_col2 = 'FourD'
    end_value_col1 = 'DilutedEarningsLossPerShareFromContinuingAndDiscontinuedOperations'
    end_value_col2 = 'FourD'
    
    start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
    if not start_index.empty:
        start_index = start_index[0]
        end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
        if not end_index.empty:
            end_index = end_index[0]
            edited_df = df.loc[start_index:end_index]
            edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
            final_data = pd.concat([general_data, edited_df], ignore_index=True)
            return final_data
    return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

def Extract_Balance_Sheet(df):
    general_data = Extract_General_Data(df)
    start_value_col1 = 'PropertyPlantAndEquipment'
    start_value_col2 = 'OneI'
    end_value_col1 = 'EquityAndLiabilities'
    end_value_col2 = 'OneI'
    
    start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
    if not start_index.empty:
        start_index = start_index[0]
        end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
        if not end_index.empty:
            end_index = end_index[0]
            edited_df = df.loc[start_index:end_index]
            edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
            final_data = pd.concat([general_data, edited_df], ignore_index=True)
            return final_data
    return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

def Extract_Equity_Adjustment(df):
    start_value_col1 = 'DescriptionOfItemThatWillNotBeReclassifiedToProfitAndLoss'  
    start_index = df[(df['Element Name'] == start_value_col1)].index
    if not start_index.empty:
        start_index = start_index[0]
        edited_df = df.loc[start_index:]
        edited_df = edited_df[['Element Name', 'Unit', 'Fact Value']]
        return edited_df
    return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])

def Extract_CFS(df):
    general_data = Extract_General_Data(df)
    start_value_col1 = 'AdjustmentsForFinanceCosts'
    start_value_col2 = 'FourD'
    end_value_col1 = 'IncreaseDecreaseInCashAndCashEquivalents'
    end_value_col2 = 'FourD'
    
    start_index = df[(df['Element Name'] == start_value_col1) & (df['Unit'] == start_value_col2)].index
    if not start_index.empty:
        start_index = start_index[0]
        end_index = df[(df['Element Name'] == end_value_col1) & (df['Unit'] == end_value_col2)].index
        if not end_index.empty:
            end_index = end_index[0]
            edited_df = df.loc[start_index:end_index]
            edited_df = edited_df[['Element Name', 'Fact Value']]
            final_data = pd.concat([general_data, edited_df], ignore_index=True)
            return final_data
    return pd.DataFrame(columns=['Element Name', 'Fact Value'])

def Extract_Segment(df):
    start_value_col1 = 'DescriptionOfReportableSegment'
    end_value_col1 = 'NetSegmentLiabilities'
    
    # Identify the segment of interest
    start_index = df[(df['Element Name'] == start_value_col1)].index
    if not start_index.empty:
        start_index = start_index[0]
        end_index = df[(df['Element Name'] == end_value_col1)].index
        if not end_index.empty:
            end_index = end_index[0]
            segment_df = df.loc[start_index:end_index]
            
            # Filter the rows where 'Unit' is 'OneD'
            oned_rows = segment_df[segment_df['Unit'] == 'OneD']
            if not oned_rows.empty:
                # Send these rows to the Extract_Income_Statement_Current function
                income_statement_df = Extract_Income_Statement_Current(oned_rows)
                
                # Remove 'OneD' rows from segment_df
                segment_df = segment_df[segment_df['Unit'] != 'OneD']
            
            # Return the remaining segment_df
            segment_df = segment_df[['Element Name', 'Unit', 'Fact Value']]
            return segment_df

    return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])


def Extract_Adjustment2(df):
    set_11 = 'AmountOfItemThatWillBeReclassifiedToProfitAndLoss'
    set_12 = 'OneD'
    set_21 = 'IncomeTaxRelatingToItmesThatWillBeReclassifiedToProfitOrLoss'
    set_22 = 'OneD'
    set_31 = 'AmountOfItemThatWillNotBeReclassifiedToProfitAndLoss'
    set_32 = 'OneD'
    set_41 = 'IncomeTaxRelatingToItmesThatWillNotBeReclassifiedToProfitAndLoss'
    set_42 = 'OneD'

    def extract_value(df, element_name, unit_name):
        filtered_df = df[(df['Element Name'] == element_name) & (df['Unit'] == unit_name)]
        if not filtered_df.empty:
            value = filtered_df['Fact Value'].iloc[0]
            try:
                return float(str(value).replace(',', ''))
            except ValueError:
                return 0
        return 0

    value_1 = extract_value(df, set_11, set_12)
    value_2 = extract_value(df, set_21, set_22)
    value_3 = extract_value(df, set_31, set_32)
    value_4 = extract_value(df, set_41, set_42)

    # Calculate the net total
    Ans1 = value_1 + value_2 - value_3 - value_4

    # Create lists for DataFrame construction
    element_names = [
        'Amount of Item That Will Be Reclassified To Profit Or Loss',
        'Amount of Item That Will Not Be Reclassified To Profit and Loss',
        'Income Tax Relating To Items That Will Not Be Reclassified To Profit Or Loss',
        'Net Total'
    ]

    fact_values = [value_1, value_2, value_3, value_4, Ans1]

    # Ensure both lists have the same length
    if len(element_names) != len(fact_values):
        # If lengths don't match, adjust the lists to have the same length
        min_length = min(len(element_names), len(fact_values))
        element_names = element_names[:min_length]
        fact_values = fact_values[:min_length]

    # Create DataFrame
    data = {
        'Element Name': element_names,
        'Fact Value': fact_values
    }

    edited_df = pd.DataFrame(data)
    return edited_df

files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]


import os
import pandas as pd

def save_results():
    function_dict = {
        'Extract_Income_Statement_Current': Extract_Income_Statement_Current,
        'Extract_Income_Statement_YTD': Extract_Income_Statement_YTD,
        'Extract_Balance_Sheet': Extract_Balance_Sheet,
        'Extract_Equity_Adjustment': Extract_Equity_Adjustment,
        'Extract_CFS': Extract_CFS,
        'Extract_Segment': Extract_Segment,
        'Extract_Adjustment2': Extract_Adjustment2
    }

    for func_name, func in function_dict.items():
        file_path = os.path.join(output_folder, f"{func_name}.xlsx")
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            sheet_data = {}  # Dictionary to hold data for each company
            
            for file in files:
                company_name = os.path.splitext(file)[0]
                df = pd.read_excel(os.path.join(folder_path, file))
                result_df = func(df)
                
                # Check if result_df is not empty before adding it
                if not result_df.empty:
                    if company_name not in sheet_data:
                        sheet_data[company_name] = []
                    sheet_data[company_name].append(result_df)
            
            # Check if there is any data to write
            if sheet_data:
                for company_name, dfs in sheet_data.items():
                    combined_df = pd.concat(dfs, ignore_index=True)
                    
                    # Ensure sheet name is valid
                    sheet_name = company_name[:31]
                    if sheet_name:  # Ensure sheet_name is not empty
                        combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"Data for {company_name} saved in {func_name}.xlsx")
                    else:
                        print(f"Skipped saving data for {company_name} due to invalid sheet name.")
            else:
                print(f"No data to save for {func_name}.xlsx")

# Run the function to save results
save_results()



def combine_company_sheets(input_dir="excel", workbook_dir="workbook", output_dir="combined"):
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Get the list of all files in the input directory
    dir_list = os.listdir(input_dir)

    # Extract company names from filenames assuming format is now "DD-MMM-YYYYCompany Name"
    company_list = list(dict.fromkeys([x[len(x.split()[0]): -5] for x in dir_list if len(x.split()) > 1]))
    
    # Get the list of workbook files
    file_list = os.listdir(workbook_dir)

    # Loop through each workbook file
    for files in file_list:
        xls = pd.ExcelFile(os.path.join(workbook_dir, files))
        print(f"Processing workbook: {files}")
        company_df = []

        # Loop through each company
        for company in company_list:
            company_sheets = []

            # Loop through each directory/file to match the company name
            for i in range(len(dir_list)):
                if company in dir_list[i]:
                    try:
                        name = dir_list[i]
                        print(f"Processing sheet: {name}")
                        # Read the Excel sheet and filter columns
                        df = pd.read_excel(xls, name)[["Element Name", "Fact Value"]]
                        df.rename(columns={"Fact Value": company}, inplace=True)
                        company_sheets.append(df)
                    except Exception as e:
                        print(f"Error reading sheet: {e}")

            # Combine all sheets for the company
            if company_sheets:
                # Determine the sheet with the maximum number of unique Element Names
                max_elements_sheet = max(company_sheets, key=lambda df: len(df['Element Name'].unique()))

                # Create a base dataframe with all unique Element Names
                base_df = max_elements_sheet[['Element Name']].drop_duplicates().reset_index(drop=True)

                # Merge all sheets based on the Element Name
                for df in company_sheets:
                    # Ensure that the sheet has the correct columns
                    df = df[['Element Name', company]]
                    base_df = pd.merge(base_df, df, on='Element Name', how='left')

                # Append the combined dataframe for this company
                company_df.append(base_df)

        # Write the combined data into a new Excel file
        try:
            output_file = os.path.join(output_dir, files)
            with pd.ExcelWriter(output_file) as writer:
                for i, df in enumerate(company_df):
                    sheet_name = company_list[i][:31]  # Truncate sheet name to 31 characters
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Workbook saved: {output_file}")
        except Exception as e:
            print(f"Error writing to Excel: {e}")

# Call the function
combine_company_sheets()
