
import pandas as pd




def load(map_filename):
    with pd.ExcelFile(map_filename,engine='openpyxl') as xls:
        overview_df = xls.parse(sheet_name='Overview', header=0, index_col=0, keep_default_na=False).fillna('')
        variables_df = xls.parse(sheet_name='Variables', header=0, index_col=0, keep_default_na=False).fillna('')
        categories_df = xls.parse(sheet_name='Analysis Values', header=0, index_col=0, keep_default_na=False).fillna('')
        validationissues_df = xls.parse(sheet_name='Validation Issues Log', header=0, index_col=0, keep_default_na=False).fillna('')
        mdd_data_variables_df = xls.parse(sheet_name='MDD_Data_Variables', header=0, index_col=0, keep_default_na=False).fillna('')
        mdd_data_categories_df = xls.parse(sheet_name='MDD_Data_Categories', header=0, index_col=0, keep_default_na=False).fillna('')
        # mdd_data_validation_df = xls.parse(sheet_name='ValidationVariables', header=0, index_col=0, keep_default_na=False).fillna('')
        result = (overview_df, variables_df, categories_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df)
        return result



