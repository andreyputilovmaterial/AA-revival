
import pandas as pd



if __name__ == '__main__':
    # run as a program
    import excel_format_sheet
    import columns_sheet_overview
    import columns_sheet_mdddata_variables
    import columns_sheet_mdddata_categories
    import columns_sheet_variables
    import columns_sheet_analysisvalues
    import columns_sheet_validation
elif '.' in __name__:
    # package
    from . import excel_format_sheet
    from . import columns_sheet_overview
    from . import columns_sheet_mdddata_variables
    from . import columns_sheet_mdddata_categories
    from . import columns_sheet_variables
    from . import columns_sheet_analysisvalues
    from . import columns_sheet_validation
else:
    # included with no parent package
    import excel_format_sheet
    import columns_sheet_overview
    import columns_sheet_mdddata_variables
    import columns_sheet_mdddata_categories
    import columns_sheet_variables
    import columns_sheet_analysisvalues
    import columns_sheet_validation



def write_to_file(dataframes,out_filename):
    overview_df, variables_df, categories_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df = dataframes
    # excel_format_sheet.format_sheet_overview(overview_df)
    # excel_format_sheet.format_sheet_variables(variables_df)
    # excel_format_sheet.format_sheet_analysisvalues(categories_df)
    with pd.ExcelWriter(out_filename, engine='openpyxl') as writer:
        overview_df.to_excel(writer, sheet_name=columns_sheet_overview.sheet_name)
        variables_df.to_excel(writer, sheet_name=columns_sheet_variables.sheet_name)
        categories_df.to_excel(writer, sheet_name=columns_sheet_analysisvalues.sheet_name)
        validationissues_df.to_excel(writer, sheet_name=columns_sheet_validation.sheet_name)
        mdd_data_variables_df.to_excel(writer, sheet_name=columns_sheet_mdddata_variables.sheet_name)
        mdd_data_categories_df.to_excel(writer, sheet_name=columns_sheet_mdddata_categories.sheet_name)
        # mdd_data_validation_df.to_excel(writer, sheet_name='ValidationVariables')
        excel_format_sheet.format_sheet_overview(writer.sheets[columns_sheet_overview.sheet_name])
        excel_format_sheet.format_with_stdstyles(writer.sheets[columns_sheet_variables.sheet_name])
        excel_format_sheet.format_sheet_variables(writer.sheets[columns_sheet_variables.sheet_name])
        excel_format_sheet.format_with_stdstyles(writer.sheets[columns_sheet_analysisvalues.sheet_name])
        excel_format_sheet.format_sheet_analysisvalues(writer.sheets[columns_sheet_analysisvalues.sheet_name])
        excel_format_sheet.format_with_stdstyles(writer.sheets[columns_sheet_validation.sheet_name])
        excel_format_sheet.format_sheet_validation(writer.sheets[columns_sheet_validation.sheet_name])
        excel_format_sheet.format_with_stdstyles(writer.sheets[columns_sheet_mdddata_variables.sheet_name])
        excel_format_sheet.format_with_stdstyles(writer.sheets[columns_sheet_mdddata_categories.sheet_name])
        # excel_format_sheet.format_with_stdstyles(writer.sheets['ValidationVariables'])
        writer.sheets[columns_sheet_mdddata_variables.sheet_name].protection.sheet = True # openpyxl-specific syntax
        writer.sheets[columns_sheet_mdddata_variables.sheet_name].sheet_state = 'hidden' # openpyxl-specific syntax
        writer.sheets[columns_sheet_mdddata_categories.sheet_name].protection.sheet = True # openpyxl-specific syntax
        writer.sheets[columns_sheet_mdddata_categories.sheet_name].sheet_state = 'hidden' # openpyxl-specific syntax
        # writer.sheets['ValidationVariables'].protection.sheet = True # openpyxl-specific syntax
        # writer.sheets['ValidationVariables'].sheet_state = 'hidden' # openpyxl-specific syntax
        # writer.sheets[columns_sheet_mdddata_variables.sheet_name].protect()
        # writer.sheets[columns_sheet_mdddata_categories.sheet_name].protect()
        # writer.sheets['ValidationVariables'].protect()
        # writer.sheets[columns_sheet_mdddata_variables.sheet_name].hide()
        # writer.sheets[columns_sheet_mdddata_categories.sheet_name].hide()
        # writer.sheets['ValidationVariables'].hide()



