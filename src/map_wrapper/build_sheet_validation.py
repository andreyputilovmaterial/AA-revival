
import pandas as pd





from . import util_dataframe_wrapper
from .column_definitions import sheet_validation
from .column_definitions import sheet_mdddata_variables
from .column_definitions import sheet_mdddata_categories



def build_df(config):
    df = util_dataframe_wrapper.PandasDataframeWrapper(sheet_validation.columns).to_df()
    df.loc['Variables'] = '=_xlfn.XLOOKUP(FALSE,{range_validation_variables_resultlookup},{range_validation_variables_namelookup},0)'.format(
            range_validation_variables_resultlookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = sheet_mdddata_variables.sheet_name,
                col = sheet_mdddata_variables.column_letters['col_validation'],
            ),
            range_validation_variables_namelookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = sheet_mdddata_variables.sheet_name,
                col = sheet_mdddata_variables.column_letters['col_question'],
            ),
    )
    df.loc['Categories'] = '=_xlfn.XLOOKUP(FALSE,{range_validation_categories_resultlookup},{range_validation_categories_namelookup},0)'.format(
            range_validation_categories_resultlookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = sheet_mdddata_categories.sheet_name,
                col = sheet_mdddata_categories.column_letters['col_validation'],
            ),
            range_validation_categories_namelookup = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = sheet_mdddata_categories.sheet_name,
                col = sheet_mdddata_categories.column_letters['col_question'],
            ),
    )
    return df



