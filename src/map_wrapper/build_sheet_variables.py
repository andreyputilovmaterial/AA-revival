

import re
import pandas as pd



from . import util_performance_monitor as util_perf
from . import util_dataframe_wrapper
from .column_definitions import sheet_variables as sheet
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
    mdmvariables = mdd.variables
    data = util_dataframe_wrapper.PandasDataframeWrapper(sheet.columns)

    performance_counter = iter(util_perf.PerformanceMonitor(config={
        'total_records': len(mdmvariables),
        'report_frequency_records_count': 1,
        'report_frequency_timeinterval': 14,
        'report_text_pipein': 'progress populating variables sheet',
    }))
    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue
        next(performance_counter)
        if not mdd.is_data_item(mdmvariable):
            continue
        if mdd.is_nocasedata(mdmvariable):
            continue
    
        substitutes = {
            **sheet.column_letters,
            'row': data.get_working_row_number(),
            'sheet_name_mdddata': sheet_mdddata_variables.sheet_name,
            'col_mdddata_question_type': '',
            # 'col_mdddata_col_last': sheet_mdddata_variables.column_letters['col_validation'],
            'col_mdddata_question_type': sheet_mdddata_variables.column_letters['col_question_type'],
            'col_mdddata_question_type_vlookup_index': sheet_mdddata_variables.column_vlookup_index['col_question_type'],
            'col_mdddata_validation': sheet_mdddata_variables.column_letters['col_validation'],
            'col_mdddata_validation_vlookup_index': sheet_mdddata_variables.column_vlookup_index['col_validation'],
        }

        question_name = mdmvariable.FullName

        question_type = '=VLOOKUP($A{row},\'{sheet_name_mdddata}\'!$A$2:${col_mdddata_question_type}$999999,{col_mdddata_question_type_vlookup_index},FALSE)&""'.format(**substitutes)

        question_shortname = prev_map.read_question_shortname(question_name)

        question_label = prev_map.read_question_label(question_name)

        question_include_truefalse = prev_map.read_question_is_included(question_name)
        
        question_include = \
            '' if question_include_truefalse=='' \
            else ( 'x' if question_include_truefalse==True or re.match(r'^\s*?(?:true)\s*?$','{s}'.format(s=question_include_truefalse),flags=re.I|re.DOTALL) \
            else ( '' if question_include_truefalse==False or re.match(r'^\s*?(?:false)\s*?$','{s}'.format(s=question_include_truefalse),flags=re.I|re.DOTALL) \
            else ( '' if has_value_numeric(question_include_truefalse) \
            else ( 'x' if question_include_truefalse else '' ) ) ) )

        question_validation_truefalse = 'VLOOKUP($A{row},\'{sheet_name_mdddata}\'!$A$2:${col_mdddata_validation}$999999,{col_mdddata_validation_vlookup_index},FALSE)'.format(**substitutes)
        question_validation = '=IF({val},"","Failed")'.format(val=question_validation_truefalse)

        question_comment = prev_map.read_question_comment(question_name)

        row_add = [
            question_name,
            question_type,
            question_label,
            question_shortname,
            question_include,
            question_validation,
            question_comment,
        ]
        data.append(*row_add)

    return data.to_df()




