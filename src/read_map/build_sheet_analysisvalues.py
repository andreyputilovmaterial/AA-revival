
import pandas as pd
from openpyxl.utils import get_column_letter



if __name__ == '__main__':
    # run as a program
    # import util_dataframe_wrapper
    import util_performance_monitor as util_perf
    import util_mdmvars
    import columns_sheet_mdddata_categories
    import columns_sheet_analysisvalues as sheet
    import columns_sheet_mdddata_variables
elif '.' in __name__:
    # package
    # from . import util_dataframe_wrapper
    from . import util_performance_monitor as util_perf
    from . import util_mdmvars
    from . import columns_sheet_mdddata_categories
    from . import columns_sheet_analysisvalues as sheet
    from . import columns_sheet_mdddata_variables
else:
    # included with no parent package
    # import util_dataframe_wrapper
    import util_performance_monitor as util_perf
    import util_mdmvars
    import columns_sheet_mdddata_categories
    import columns_sheet_analysisvalues as sheet
    import columns_sheet_mdddata_variables



def build_df(mdmvariables,df_prev,config):
    # data = util_dataframe_wrapper.PandasDataframeWrapper(sheet.columns)
    data_add = []
    row = 2

    performance_counter = iter(util_perf.PerformanceMonitor(config={
        'total_records': len(mdmvariables),
        'report_frequency_records_count': 1,
        'report_frequency_timeinterval': 14,
        'report_text_pipein': 'building analysis values dataframe',
    }))

    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue
        category_question_name = mdmvariable.FullName
        next(performance_counter)
        if util_mdmvars.is_iterative(mdmvariable) or util_mdmvars.has_own_categories(mdmvariable):


            substitutes = {
                **sheet.column_letters,
                # 'row': data.get_working_row_number(),
                'row': row,
                'sheet_name_mdddata_categories': columns_sheet_mdddata_categories.sheet_name,
                'sheet_name_mdddata_variables': columns_sheet_mdddata_variables.sheet_name,
                'col_mddvars_question': columns_sheet_mdddata_variables.column_letters['col_question'],
                'col_mddvars_shortname': columns_sheet_mdddata_variables.column_letters['col_shortname'],
                'col_mddvars_shortname_vlookup_index': columns_sheet_mdddata_variables.column_vlookup_index['col_shortname'],
                'col_mddvars_include': columns_sheet_mdddata_variables.column_letters['col_include'],
                'col_mddvars_include_vlookup_index': columns_sheet_mdddata_variables.column_vlookup_index['col_include'],
                'col_mddcats_path': columns_sheet_mdddata_categories.column_letters['col_path'],
                'col_mddcats_validation': columns_sheet_mdddata_categories.column_letters['col_validation'],
                'col_mddcats_validation_vlookup_index': columns_sheet_mdddata_categories.column_vlookup_index['col_validation'],
            }

            categories_data = []
            category_question_shortname       = '=VLOOKUP(${col_question}{row},\'{sheet_name_mdddata_variables}\'!${col_mddvars_question}$2:${col_mddvars_shortname}$999999,{col_mddvars_shortname_vlookup_index},FALSE)'.format(**substitutes)
            category_question_include_truefalse = 'VLOOKUP(${col_question}{row},\'{sheet_name_mdddata_variables}\'!${col_mddvars_question}$2:${col_mddvars_include}$999999,{col_mddvars_include_vlookup_index},FALSE)'.format(**substitutes)
            category_question_in_exclusions   = '=IF({val},"","(exclude)")'.format(val=category_question_include_truefalse)

            for col_index_zerobased, mdmcategory in enumerate(util_mdmvars.list_mdmcategories(mdmvariable)):

                col_index = 3 + col_index_zerobased
                col_letter = get_column_letter(col_index)
                category_name = mdmcategory.Name
                category_path = '{var}.Categories[{cat}]'.format(var=category_question_name,cat=category_name)

                substitutes['col'] = col_letter

                category_path_as_formula = '""&$A{row}&".Categories["&{col}{row}&"]"'.format(**substitutes)

                substitutes['category_path_as_formula'] = category_path_as_formula

                category_label = \
                    ( \
                        df_prev.loc[category_path,columns_sheet_mdddata_categories.column_names['col_label_prev']] \
                        if df_prev.loc[category_path,columns_sheet_mdddata_categories.column_names['col_label_prev']] \
                        else df_prev.loc[category_path,columns_sheet_mdddata_categories.column_names['col_label_mdd']] \
                    ) \
                    if df_prev is not None and category_path in df_prev.index.get_level_values(0) \
                    else ''

                category_validation = '=VLOOKUP({category_path_as_formula},\'{sheet_name_mdddata_categories}\'!${col_mddcats_path}$2:${col_mddcats_validation}$999999,{col_mddcats_validation_vlookup_index},FALSE)'.format(**substitutes)
                
                category_analysisvalue = \
                    ( \
                        df_prev.loc[category_path,columns_sheet_mdddata_categories.column_names['col_analysisvalue_prev']] \
                        if df_prev.loc[category_path,columns_sheet_mdddata_categories.column_names['col_analysisvalue_prev']] \
                        else df_prev.loc[category_path,columns_sheet_mdddata_categories.column_names['col_analysisvalue_mdd']] \
                    ) \
                    if df_prev is not None and category_path in df_prev.index.get_level_values(0) \
                    else ''

                categories_data.append({
                    'name': category_name,
                    'label': category_label,
                    'value': category_analysisvalue,
                    'validation': category_validation,
                })
            
            data_add.append({
                'Question':category_question_name,
                'ShortName': category_question_shortname,
                'is_excluded': category_question_in_exclusions,
                'data': categories_data,
            })
            row = row + 4

    cat_columns_count = max([len(record['data']) for record in data_add])
    # df = util_dataframe_wrapper.PandasDataframeWrapper(sheet.columns).to_df()
    df =  df = pd.DataFrame(
        data =  { col[0]: col[1] for col in zip(sheet.columns,[[]]*len(sheet.columns)) },
        index = None,
    )
    df = pd.concat(
        [
            df,
            pd.DataFrame(
                [],
                index=df.index,
                columns = [ 'CAT_{n}'.format(n=0+n+1) for n in range(0,cat_columns_count) ]
            )
        ],
        axis = 1
    )
    
    data_add_ready = []
    for record in data_add:

        # 1. a row with category name
        row_add = [
            record['Question'],
            record['ShortName'],
        ] \
        + [
            cell['name'] for cell in record['data']
        ] \
        + [''] * ( cat_columns_count - len(record['data']) )
        data_add_ready.append(row_add)

        # 2. a row with category label
        row_add = [
            '{var} : Label'.format(var=record['Question']),
            '',
        ] \
        + [
            cell['label'] for cell in record['data']
        ] \
        + [''] * ( cat_columns_count - len(record['data']) )
        data_add_ready.append(row_add)

        # 3. a row with analysis values
        row_add = [
            '{var} : Analysis Value'.format(var=record['Question']),
            record['is_excluded'],
        ] \
        + [
            cell['value'] for cell in record['data']
        ] \
        + [''] * ( cat_columns_count - len(record['data']) )
        data_add_ready.append(row_add)

        # 4. a row with category label
        row_add = [
            '{var} : Validation'.format(var=record['Question']),
            '',
        ] \
        + [
            cell['validation'] for cell in record['data']
        ] \
        + [''] * ( cat_columns_count - len(record['data']) )
        data_add_ready.append(row_add)

    data_add_ready = [
        { col[0]: col[1] for col in zip(df.columns,record) } for record in data_add_ready
    ]
    data_add = data_add_ready

    df = pd.concat([df,pd.DataFrame(data_add)],ignore_index=True)
    df.set_index('Question',inplace=True)
    return df

