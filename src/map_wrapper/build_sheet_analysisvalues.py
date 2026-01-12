
import pandas as pd
from openpyxl.utils import get_column_letter



# from . import util_dataframe_wrapper
from . import util_performance_monitor as util_perf
from .column_definitions import sheet_mdddata_categories
from .column_definitions import sheet_analysisvalues as sheet
from .column_definitions import sheet_mdddata_variables


def has_value_numeric(arg):
    if pd.isna(arg):
        return False
    if arg is None:
        return False
    if arg==0:
        return True
    if arg == False:
        return True # false evaluates to 0 which is numeric
    if arg=='':
        return False
    return not not arg

def has_value_text(arg):
    if pd.isna(arg):
        return False
    if arg is None:
        return False
    if arg==0:
        return True
    if arg == False:
        return False
    if arg=='':
        return False
    return not not arg



def build_df(mdd,prev_map,config):
    # data = util_dataframe_wrapper.PandasDataframeWrapper(sheet.columns)
    mdmvariables = mdd.variables
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
        if mdd.is_iterative(mdmvariable) or mdd.has_own_categories(mdmvariable):

            try:

                substitutes = {
                    **sheet.column_letters,
                    # 'row': data.get_working_row_number(),
                    'row': row,
                    'sheet_name_mdddata_categories': sheet_mdddata_categories.sheet_name,
                    'sheet_name_mdddata_variables': sheet_mdddata_variables.sheet_name,
                    'col_mddvars_question': sheet_mdddata_variables.column_letters['col_question'],
                    'col_mddvars_shortname': sheet_mdddata_variables.column_letters['col_shortname'],
                    'col_mddvars_shortname_vlookup_index': sheet_mdddata_variables.column_vlookup_index['col_shortname'],
                    'col_mddvars_include': sheet_mdddata_variables.column_letters['col_include'],
                    'col_mddvars_include_vlookup_index': sheet_mdddata_variables.column_vlookup_index['col_include'],
                    'col_mddcats_path': sheet_mdddata_categories.column_letters['col_path'],
                    'col_mddcats_validation': sheet_mdddata_categories.column_letters['col_validation'],
                    'col_mddcats_validation_vlookup_index': sheet_mdddata_categories.column_vlookup_index['col_validation'],
                }

                categories_data = []
                category_question_shortname       = '=VLOOKUP(${col_question}{row},\'{sheet_name_mdddata_variables}\'!${col_mddvars_question}$2:${col_mddvars_shortname}$999999,{col_mddvars_shortname_vlookup_index},FALSE)'.format(**substitutes)
                category_question_include_truefalse = 'VLOOKUP(${col_question}{row},\'{sheet_name_mdddata_variables}\'!${col_mddvars_question}$2:${col_mddvars_include}$999999,{col_mddvars_include_vlookup_index},FALSE)'.format(**substitutes)
                category_question_in_exclusions   = '=IF({val},"","(exclude)")'.format(val=category_question_include_truefalse)

                for col_index_zerobased, mdmcategory in enumerate(mdd.list_mdmcategories(mdmvariable)):

                    try:
                        
                        col_index = 3 + col_index_zerobased
                        col_letter = get_column_letter(col_index)
                        category_name = mdmcategory.Name
                        # category_path = '{var}.Categories[{cat}]'.format(var=category_question_name,cat=category_name)

                        substitutes['col'] = col_letter

                        category_path_as_formula = '""&$A{row}&".Categories["&{col}{row}&"]"'.format(**substitutes)

                        substitutes['category_path_as_formula'] = category_path_as_formula

                        category_label = prev_map.read_category_label(category_question_name,category_name)
                        if not prev_map.has_value_text(category_label):
                            category_label = prev_map.read_category_label_column_mdd(category_question_name,category_name)

                        category_validation = '=VLOOKUP({category_path_as_formula},\'{sheet_name_mdddata_categories}\'!${col_mddcats_path}$2:${col_mddcats_validation}$999999,{col_mddcats_validation_vlookup_index},FALSE)'.format(**substitutes)
                        
                        category_analysisvalue = prev_map.read_category_analysisvalue(category_question_name,category_name)
                        if not prev_map.has_value_numeric(category_analysisvalue):
                            category_analysisvalue = prev_map.read_category_analysisvalue_column_mdd(category_question_name,category_name)

                        categories_data.append({
                            'name': category_name,
                            'label': category_label,
                            'value': category_analysisvalue,
                            'validation': category_validation,
                        })
                    
                    except Exception as e:
                        print('Failed when processing category {cat}, Error: {e}'.format(cat=mdmcategory.Name,e=e)) # TODO: print to stderr
                        raise e
                        
                data_add.append({
                    'Question':category_question_name,
                    'ShortName': category_question_shortname,
                    'is_excluded': category_question_in_exclusions,
                    'data': categories_data,
                })
                row = row + 4

            except Exception as e:
                print('Failed when processing variable {mdmvar}, Error: {e}'.format(mdmvar=mdmvariable.Name,e=e))
                raise e


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



def prepopulate(df_sheet_analysisvalues,df_sheet_categories,config):
    df_sheet_categories = df_sheet_categories.copy()
    variables_to_process = []
    for rownumber in range(0,len(df_sheet_categories)):
        pass
    return df_sheet_analysisvalues