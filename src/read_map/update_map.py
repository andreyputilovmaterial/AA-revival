
import pandas as pd





if __name__ == '__main__':
    # run as a program
    import build_sheet_overview
    import build_sheet_variables
    import build_sheet_analysisvalues
    import build_sheet_mdddata_variables
    import build_sheet_mdddata_categories
    import build_sheet_validation
elif '.' in __name__:
    # package
    from . import build_sheet_overview
    from . import build_sheet_variables
    from . import build_sheet_analysisvalues
    from . import build_sheet_mdddata_variables
    from . import build_sheet_mdddata_categories
    from . import build_sheet_validation
else:
    # included with no parent package
    import build_sheet_overview
    import build_sheet_variables
    import build_sheet_analysisvalues
    import build_sheet_mdddata_variables
    import build_sheet_mdddata_categories
    import build_sheet_validation






def update_map(mdd_variables,dataframes,config):
    assert mdd_variables, 'update_map: mdd is required but it missing'
    if not dataframes:
        dataframes = (
            None,
            None,
            None,
            None,
            None,
            None,
        )

    overview_df, variables_df, categories_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df = dataframes

    print('building "Overview" sheet...')
    overview_df                 = build_sheet_overview.build_df(mdd_variables,overview_df,config)
    print('building "MDD_Data_Variables" (hidden) sheet...')
    mdd_data_variables_df       = build_sheet_mdddata_variables.build_df(mdd_variables,mdd_data_variables_df,config)
    print('building "Variables" sheet...')
    variables_df                = build_sheet_variables.build_df(mdd_variables,mdd_data_variables_df,config)
    print('building "MDD_Data_Categories" (hidden) sheet...')
    mdd_data_categories_df      = build_sheet_mdddata_categories.build_df(mdd_variables,mdd_data_categories_df,config)
    print('building "Analysis Values" sheet...')
    categories_df               = build_sheet_analysisvalues.build_df(mdd_variables,mdd_data_categories_df,config)
    print('building "Validation Issues Log" sheet...')
    validationissues_df         = build_sheet_validation.build_df(mdd_variables,validationissues_df,config)

    result = (overview_df, variables_df, categories_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df)
    return result



