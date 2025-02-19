
import pandas as pd

from openpyxl.utils import get_column_letter



if __name__ == '__main__':
    # run as a program
    import util_performance_monitor as util_perf
    import util_mdmvars
elif '.' in __name__:
    # package
    from . import util_performance_monitor as util_perf
    from . import util_mdmvars
else:
    # included with no parent package
    import util_performance_monitor as util_perf
    import util_mdmvars



def build_df(mdmvariables,df_prev,config):
    df = pd.DataFrame(
        data = {
            'Question': [],
            'ShortName': [],
        },
        index = None,
    )
    return df

    # adding to dataframe row by row is very slow
    # that's why
    # I set key to None,
    # I am adding all to a list with dicts
    # and then concatenate the df with it
    # and set index back, this is much faster
    # (do we even need index here?)
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

            categories_data = []
            category_question_shortname = '=VLOOKUP($A{row},MDD_Data_Variables!$A$2:$H$699999,5,FALSE)'.format(row=row)
            # category_question_in_exclusions = '=IF(VLOOKUP($A{row},Variables!$A$2:$G$699999,6,FALSE)&""="","(excluded)","")'.format(row=row)
            # category_question_in_exclusions = '=VLOOKUP($A{row},MDD_Data_Variables!$A$2:$H$699999,7,FALSE)'.format(row=row)
            category_question_in_exclusions = '=IF(VLOOKUP($A{row},MDD_Data_Variables!$A$2:$H$699999,7,FALSE)="","(in exclusions)",IF(VLOOKUP($A{row},MDD_Data_Variables!$A$2:$H$699999,7,FALSE)="x","",VLOOKUP($A{row},MDD_Data_Variables!$A$2:$H$699999,7,FALSE)))'.format(row=row)

            for col_index_zerobased, mdmcategory in enumerate(util_mdmvars.list_mdmcategories(mdmvariable)):

                col_index = 3+col_index_zerobased
                col_letter = get_column_letter(col_index)
                category_name = mdmcategory.Name
                category_item_name ='{var}.Categories[{cat}]'.format(var=category_question_name,cat=category_name)
                category_label = '=VLOOKUP(""&$A{row}&".Categories["&{col}{row}&"]",MDD_Data_Categories!$A$2:$G$699999,5,FALSE)'.format(row=row,col=col_letter)
                category_validation = '=VLOOKUP(""&$A{row}&".Categories["&{col}{row}&"]",MDD_Data_Categories!$A$2:$G$699999,7,FALSE)'.format(row=row,col=col_letter)
                
                category_analysisvalue = None
                try:
                    if category_item_name in mdd_data_categories_prev_df.index.get_level_values(0):
                        category_analysisvalue = aa_logic.sanitize_analysis_value(mdd_data_categories_prev_df.loc[category_item_name,'AnalysisValue'])
                except Exception as e:
                    raise e
                    # pass
                if category_analysisvalue is None:
                    category_analysisvalue = aa_logic.sanitize_analysis_value(mdmcategory.Properties['Value'])
                if category_analysisvalue is None:
                    category_analysisvalue = ''

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

