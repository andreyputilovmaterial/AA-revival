
import pandas as pd




if __name__ == '__main__':
    # run as a program
    import util_dataframe_wrapper
    import columns_sheet_validation
    import columns_sheet_mdddata_variables
    import columns_sheet_mdddata_categories
elif '.' in __name__:
    # package
    from . import util_dataframe_wrapper
    from . import columns_sheet_validation
    from . import columns_sheet_mdddata_variables
    from . import columns_sheet_mdddata_categories
else:
    # included with no parent package
    import util_dataframe_wrapper
    import columns_sheet_validation
    import columns_sheet_mdddata_variables
    import columns_sheet_mdddata_categories



def build_df(variables,df_prev,config):
    df = util_dataframe_wrapper.PandasDataframeWrapper(columns_sheet_validation.columns).to_df()
    df.loc['Variables'] = '=_xlfn.XLOOKUP(FALSE,{range_validation_variables_resultlookup},{range_validation_variables_namelookup},0)'.format(
            range_validation_variables_resultlookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = columns_sheet_mdddata_variables.sheet_name,
                col = columns_sheet_mdddata_variables.column_letters['col_validation'],
            ),
            range_validation_variables_namelookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = columns_sheet_mdddata_variables.sheet_name,
                col = columns_sheet_mdddata_variables.column_letters['col_question'],
            ),
    )
    df.loc['Categories'] = '=_xlfn.XLOOKUP(FALSE,{range_validation_categories_resultlookup},{range_validation_categories_namelookup},0)'.format(
            range_validation_categories_resultlookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = columns_sheet_mdddata_categories.sheet_name,
                col = columns_sheet_mdddata_categories.column_letters['col_validation'],
            ),
            range_validation_categories_namelookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = columns_sheet_mdddata_categories.sheet_name,
                col = columns_sheet_mdddata_categories.column_letters['col_question'],
            ),
    )
    return df



