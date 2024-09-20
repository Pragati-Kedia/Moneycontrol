# # code for merging the data uccefully 
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
#             print("General data extracted:", general_data)
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
#             for file in files:
#                 company_name = os.path.splitext(file)[0]
#                 df = pd.read_excel(os.path.join(folder_path, file))
#                 result_df = func(df)
#                 if not result_df.empty:
#                     if company_name not in sheet_data:
#                         sheet_data[company_name] = []
#                     sheet_data[company_name].append(result_df)
            
#             # Write all accumulated data to a single sheet for each company
#             for company_name, dfs in sheet_data.items():
#                 combined_df = pd.concat(dfs, ignore_index=True)
                
#                 # Truncate worksheet names to 31 characters or fewer
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
    
#     # Create a list of company names based on the filenames
#     formatted_list = [x[:31] for x in dir_list]
#     company_list = list(dict.fromkeys([x[(x.rfind("_"))+1:-5] for x in dir_list]))
    
#     # Get the list of workbook files
#     file_list = os.listdir(workbook_dir)
    
#     # Loop through each workbook file
#     for files in file_list:
#         xls = pd.ExcelFile(os.path.join(workbook_dir, files))
#         print(f"Processing workbook: {files}")
#         company_df = []
        
#         # Loop through each company
#         for company in company_list:
#             company_sheets = []
            
#             # Loop through each directory/file to match the company name
#             for i in range(len(dir_list)):
#                 if dir_list[i].find(company) > 0:
#                     try:
#                         name = formatted_list[i]
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
#             output_file = os.path.join(output_dir, files)
#             with pd.ExcelWriter(output_file) as writer:
#                 for i, df in enumerate(company_df):
#                     sheet_name = company_list[i][:31]  # Truncate sheet name to 31 characters
#                     df.to_excel(writer, sheet_name=sheet_name, index=False)
#             print(f"Workbook saved: {output_file}")
#         except Exception as e:
#             print(f"Error writing to Excel: {e}")

# # Call the function
# combine_company_sheets()


# data is not being showed in the workbooks 
# import os
# import pandas as pd
# import openpyxl

# # Define paths
# folder_path = r"D:\excel_workbook\excel"
# output_folder = r"D:\excel_workbook\workbook"
# combined_output_folder = r"D:\excel_workbook\combined"

# # Ensure output folders exist
# os.makedirs(output_folder, exist_ok=True)
# os.makedirs(combined_output_folder, exist_ok=True)

# # List of all files in the folder
# files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

# def ensure_visible_sheets(file_path):
#     """Ensure all sheets in the workbook are visible."""
#     wb = openpyxl.load_workbook(file_path)
#     for sheet in wb.sheetnames:
#         ws = wb[sheet]
#         ws.sheet_state = 'visible'
#     wb.save(file_path)

# def extract_company_name(filename):
#     """Extract company name from filename with a date prefix."""
#     name_without_ext = os.path.splitext(filename)[0]
#     for i in range(len(name_without_ext)):
#         if name_without_ext[i].isalpha():
#             return name_without_ext[i:].strip()
#     return name_without_ext

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
#             print("General data extracted:", general_data)
#             return general_data
#     return pd.DataFrame(columns=['Element Name', 'Unit', 'Fact Value'])


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
#             print("Income Statement Current data extracted:", final_data)
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
#     start_index = df[(df['Element Name'] == start_value_col1)].index
#     if not start_index.empty:
#         start_index = start_index[0]
#         end_index = df[(df['Element Name'] == end_value_col1)].index
#         if not end_index.empty:
#             end_index = end_index[0]
#             segment_df = df.loc[start_index:end_index]
#             oned_rows = segment_df[segment_df['Unit'] == 'OneD']
#             if not oned_rows.empty:
#                 income_statement_df = Extract_Income_Statement_Current(oned_rows)
#                 segment_df = segment_df[segment_df['Unit'] != 'OneD']
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

#     Ans1 = value_1 + value_2 - value_3 - value_4

#     element_names = [
#         'Amount of Item That Will Be Reclassified To Profit Or Loss',
#         'Amount of Item That Will Not Be Reclassified To Profit and Loss',
#         'Income Tax Relating To Items That Will Not Be Reclassified To Profit Or Loss',
#         'Net Total'
#     ]

#     fact_values = [value_1, value_2, value_3, Ans1]

#     adjustment_data = pd.DataFrame({
#         'Element Name': element_names,
#         'Fact Value': fact_values
#     })

#     return adjustment_data


# os.makedirs(output_folder, exist_ok=True)

# def extract_company_name(filename):
#     """Extract company name from filename with a date prefix."""
#     name_without_ext = os.path.splitext(filename)[0]
#     for i in range(len(name_without_ext)):
#         if name_without_ext[i].isalpha():
#             return name_without_ext[i:].strip()
#     return name_without_ext

# def save_results():
#     print("Starting save_results function...")
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
#         print(f"Processing function: {func_name}")
#         file_path = os.path.join(output_folder, f"{func_name}.xlsx")
#         try:
#             with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
#                 sheet_data = {}
#                 files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
#                 for file in files:
#                     print(f"Processing file: {file}")
#                     company_name = extract_company_name(file)
#                     df = pd.read_excel(os.path.join(folder_path, file))
#                     result_df = func(df)
#                     if not result_df.empty:
#                         if company_name not in sheet_data:
#                             sheet_data[company_name] = []
#                         sheet_data[company_name].append(result_df)
                
#                 if not sheet_data:
#                     dummy_df = pd.DataFrame({"Element Name": [], "Fact Value": []})
#                     dummy_df.to_excel(writer, sheet_name='NoData', index=False)
#                     print(f"No valid data found for {func_name}. Added dummy sheet.")
#                 else:
#                     for company_name, dfs in sheet_data.items():
#                         combined_df = pd.concat(dfs, ignore_index=True)
#                         sheet_name = company_name[:31]
#                         combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
#                         print(f"Data for {company_name} saved in {func_name}.xlsx")
#         except Exception as e:
#             print(f"Error writing results for {func_name}: {e}")

# # Test basic functionality
# print("Script started.")
# try:
#     save_results()
# except Exception as e:
#     print(f"Script error: {e}")

# def combine_company_sheets(input_dir=folder_path, workbook_dir=output_folder, output_dir=combined_output_folder):
#     print("Starting combine_company_sheets function...")
#     os.makedirs(output_dir, exist_ok=True)
#     dir_list = os.listdir(input_dir)
#     company_list = list(set(extract_company_name(f) for f in dir_list))
#     print(f"Extracted company names: {company_list}")

#     file_list = os.listdir(workbook_dir)
#     for file in file_list:
#         print(f"Processing workbook: {file}")
#         xls_path = os.path.join(workbook_dir, file)
#         ensure_visible_sheets(xls_path)
#         try:
#             xls = pd.ExcelFile(xls_path)
#             company_df = []

#             for company in company_list:
#                 company_sheets = []
#                 for i in range(len(dir_list)):
#                     if company in extract_company_name(dir_list[i]):
#                         try:
#                             name = extract_company_name(dir_list[i])
#                             print(f"Processing sheet: {name}")
#                             if name in xls.sheet_names:
#                                 df = pd.read_excel(xls, name)[["Element Name", "Fact Value"]]
#                                 company_sheets.append(df)
#                             else:
#                                 print(f"Sheet '{name}' not found in workbook.")
#                         except Exception as e:
#                             print(f"Error reading sheet: {e}")
                
#                 if company_sheets:
#                     max_elements_sheet = max(company_sheets, key=lambda df: len(df['Element Name'].unique()))
#                     base_df = max_elements_sheet[['Element Name']].drop_duplicates().reset_index(drop=True)
                    
#                     for df in company_sheets:
#                         df = df[['Element Name', 'Fact Value']]
#                         base_df = pd.merge(base_df, df, on='Element Name', how='left')
                    
#                     company_df.append(base_df)
            
#             try:
#                 output_file = os.path.join(output_dir, file)
#                 with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#                     if company_df:
#                         for i, df in enumerate(company_df):
#                             sheet_name = company_list[i][:31]
#                             df.to_excel(writer, sheet_name=sheet_name, index=False)
#                             print(f"Written data to sheet: {sheet_name}")
#                     else:
#                         dummy_df = pd.DataFrame({"Element Name": [], "Fact Value": []})
#                         dummy_df.to_excel(writer, sheet_name='Sheet1', index=False)
#                         print("No valid sheets found. Added dummy sheet.")
#                 print(f"Workbook saved: {output_file}")
#             except Exception as e:
#                 print(f"Error writing to Excel: {e}")
#         except Exception as e:
#             print(f"Error processing workbook {file}: {e}")

# Run combine function if needed
# combine_company_sheets()


# code for merging the data uccefully for file name  Example :- 2020-2021_Q2_ACC LTD 
import os
import pandas as pd
import re


# Define paths
folder_path = r"D:\excel_workbook\excel"
output_folder = r"D:\excel_workbook\workbook"
os.makedirs(output_folder, exist_ok=True)

# List of all files in the folder
files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]

def get_company_name(filename):
    """
    Extract company name from the filename assuming the format is:
    <date> <Company Name>.xlsx or <date> <Company Name>.xls
    For example: '30-Jun-2019 ICICI Bank Limited.xlsx'
    """
    # Remove file extension
    base_name = os.path.splitext(filename)[0]

    # Regular expression to match and extract the company name
    # This assumes the date is followed by a space and then the company name
    match = re.match(r'^\d{2}-\w{3}-\d{4}\s+(.*)', base_name)

    if match:
        company_name = match.group(1).strip()
        return company_name
    else:
        # If the pattern doesn't match, return the base name as is
        return base_name


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

        # Check if file already exists and remove it to avoid conflicts
        if os.path.exists(file_path):
            os.remove(file_path)

        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            sheet_data = {}  # Dictionary to hold data for each company

            for file in files:
                company_name = get_company_name(file)
                df = pd.read_excel(os.path.join(folder_path, file))
                result_df = func(df)

                if company_name not in sheet_data:
                    sheet_data[company_name] = []
                if not result_df.empty:
                    sheet_data[company_name].append(result_df)

            # Write all accumulated data to a single sheet for each company
            for company_name, dfs in sheet_data.items():
                combined_df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

                # Ensure combined_df is not empty before saving
                sheet_name = company_name[:31]
                combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"Data for {company_name} saved in {func_name}.xlsx")

            # Ensure the workbook has at least one empty sheet if no data
            if not writer.sheets:
                pd.DataFrame().to_excel(writer, sheet_name="Blank_Sheet", index=False)
                print(f"No data found for {func_name}.xlsx, blank sheet created.")

# Run the function to save results
save_results()

def combine_company_sheets(input_dir="workbook", output_dir="combined"):
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Get the list of workbook files
    workbook_files = os.listdir(input_dir)
    
    # Loop through each workbook file
    for workbook_file in workbook_files:
        print(f"Processing workbook: {workbook_file}")
        
        # Load the workbook to inspect its sheets
        xls = pd.ExcelFile(os.path.join(input_dir, workbook_file))
        company_data = {}

        # Loop through each sheet in the workbook
        for sheet in xls.sheet_names:
            try:
                # Extract the company name from the sheet name
                company_name = sheet.split('_')[-1].strip()[:31]
                
                # Read the sheet and filter only the necessary columns
                df = pd.read_excel(xls, sheet_name=sheet)[["Element Name", "Fact Value"]]
                
                # Rename the Fact Value column to reflect the sheet (quarter)
                quarter_name = sheet[:sheet.rfind("_")]
                df.rename(columns={"Fact Value": quarter_name}, inplace=True)
                
                # Append or merge the data for this company
                if company_name not in company_data:
                    company_data[company_name] = df
                else:
                    # Merge data based on 'Element Name'
                    company_data[company_name] = pd.merge(
                        company_data[company_name], df, on="Element Name", how="left"
                    )
                
                print(f"Processed sheet: {sheet} for company: {company_name}")
            
            except Exception as e:
                print(f"Error processing sheet: {sheet} in workbook: {workbook_file}: {e}")
        
        # Save the combined data into the respective output workbook
        output_file = os.path.join(output_dir, workbook_file)
        try:
            # Save the combined company sheets back into the same workbook structure
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for company_name, df in company_data.items():
                    df.to_excel(writer, sheet_name=company_name[:31], index=False)
            print(f"Workbook saved: {output_file}")
        except Exception as e:
            print(f"Error writing to Excel for {workbook_file}: {e}")

# Call the function
combine_company_sheets()
