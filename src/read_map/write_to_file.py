
import pandas as pd




if __name__ == '__main__':
    # run as a program
    import excel_sheet_format
elif '.' in __name__:
    # package
    from . import excel_sheet_format
else:
    # included with no parent package
    import excel_sheet_format




def write_to_file(dataframes,out_filename):
    overview_df, variables_df, categories_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df = dataframes
    # excel_sheet_format.format_overview(overview_df)
    # excel_sheet_format.format_variables(variables_df)
    # excel_sheet_format.format_categories(categories_df)
    with pd.ExcelWriter(out_filename, engine='openpyxl') as writer:
        overview_df.to_excel(writer, sheet_name='Overview')
        variables_df.to_excel(writer, sheet_name='Variables')
        categories_df.to_excel(writer, sheet_name='Analysis Values')
        validationissues_df.to_excel(writer, sheet_name='Validation Issues Log')
        mdd_data_variables_df.to_excel(writer, sheet_name='MDD_Data_Variables')
        mdd_data_categories_df.to_excel(writer, sheet_name='MDD_Data_Categories')
        # mdd_data_validation_df.to_excel(writer, sheet_name='ValidationVariables')
        excel_sheet_format.format_overview(writer.sheets['Overview'])
        excel_sheet_format.format_stdstyles(writer.sheets['Variables'])
        excel_sheet_format.format_variables(writer.sheets['Variables'])
        excel_sheet_format.format_stdstyles(writer.sheets['Analysis Values'])
        excel_sheet_format.format_categories(writer.sheets['Analysis Values'])
        excel_sheet_format.format_stdstyles(writer.sheets['Validation Issues Log'])
        excel_sheet_format.format_validationissues(writer.sheets['Validation Issues Log'])
        excel_sheet_format.format_stdstyles(writer.sheets['MDD_Data_Variables'])
        excel_sheet_format.format_stdstyles(writer.sheets['MDD_Data_Categories'])
        # excel_sheet_format.format_stdstyles(writer.sheets['ValidationVariables'])
        writer.sheets['MDD_Data_Variables'].protection.sheet = True # openpyxl-specific syntax
        writer.sheets['MDD_Data_Variables'].sheet_state = 'hidden' # openpyxl-specific syntax
        writer.sheets['MDD_Data_Categories'].protection.sheet = True # openpyxl-specific syntax
        writer.sheets['MDD_Data_Categories'].sheet_state = 'hidden' # openpyxl-specific syntax
        # writer.sheets['ValidationVariables'].protection.sheet = True # openpyxl-specific syntax
        # writer.sheets['ValidationVariables'].sheet_state = 'hidden' # openpyxl-specific syntax
        # writer.sheets['MDD_Data_Variables'].protect()
        # writer.sheets['MDD_Data_Categories'].protect()
        # writer.sheets['ValidationVariables'].protect()
        # writer.sheets['MDD_Data_Variables'].hide()
        # writer.sheets['MDD_Data_Categories'].hide()
        # writer.sheets['ValidationVariables'].hide()



