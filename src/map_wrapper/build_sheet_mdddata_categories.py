

from . import aa_logic
from . import util_performance_monitor as util_perf
from . import util_dataframe_wrapper
from .column_definitions import sheet_mdddata_categories as sheet
from .column_definitions import sheet_analysisvalues
from .column_definitions import sheet_mdddata_variables



# from openpyxl.worksheet.formula import ArrayFormula
def build_df(mdd,prev_map,config):
    mdmvariables = mdd.variables
    data = util_dataframe_wrapper.PandasDataframeWrapper(sheet.columns)

    performance_counter = iter(util_perf.PerformanceMonitor(config={
        'total_records': len(mdmvariables),
        'report_frequency_records_count': 1,
        'report_frequency_timeinterval': 14,
        'report_text_pipein': 'progress reading analysis values and populating categories sheet',
    }))
    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue

        category_question_name = mdmvariable.FullName

        next(performance_counter)
        if mdd.is_iterative(mdmvariable) or mdd.has_own_categories(mdmvariable):
            for mdmsharedlist_name, mdmcategory in mdd.list_mdmcategories_with_slname(mdmvariable):

                substitutes = {
                    **sheet.column_letters,
                    'row': data.get_working_row_number(),
                    'sheet_name_analysisvalues': sheet_analysisvalues.sheet_name,
                    'sheet_analysisvalues_values_range': '${col_data_start}$2:$ZZ$999999'.format(col_data_start='C'),
                    'sheet_analysisvalues_allsheet_range': '${col_data_start}$2:$ZZ$999999'.format(col_data_start='A'),
                    'sheet_analysisvalues_index_column_range': '${col_data_index}$2:${col_data_index}$999999'.format(col_data_index='A'),
                    'sheet_name_mdddata_variables': sheet_mdddata_variables.sheet_name,
                    'col_mdddata_variables_include': sheet_mdddata_variables.column_letters['col_include'],
                    'col_mdddata_variables_include_vlookup_index': sheet_mdddata_variables.column_vlookup_index['col_include'],
                    # 'col_mdddata_question_type': '',
                    # # 'col_mdddata_col_last': sheet_analysisvalues.column_letters['col_validation'],
                    # 'col_mdddata_question_type': sheet_analysisvalues.column_letters['col_question_type'],
                    # 'col_mdddata_question_type_vlookup_index': sheet_analysisvalues.column_vlookup_index['col_question_type'],
                    # 'col_mdddata_validation': sheet_analysisvalues.column_letters['col_validation'],
                    # 'col_mdddata_validation_vlookup_index': sheet_analysisvalues.column_vlookup_index['col_validation'],
                }
                
                category_name = mdmcategory.Name
                category_path = '=""&${col_question}{row}&".Categories["&${col_category}{row}&"]"'.format(**substitutes)
                category_path = '{question_name}.Categories[{category_name}]'.format(question_name=category_question_name,category_name=category_name)

                category_sharedlist = mdmsharedlist_name if mdmsharedlist_name else ''

                category_analysisvalue_lookup_categorynames_array = '_xlfn.XLOOKUP( ${col_question}{row}&"", \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_index_column_range}, \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_values_range} )'.format(**substitutes)
                category_analysisvalue_lookup_analysisvalues_array = '_xlfn.XLOOKUP( ${col_question}{row}&" : Analysis Value", \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_index_column_range}, \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_values_range} )'.format(**substitutes)
                category_analysisvalue_lookup_labels_array = '_xlfn.XLOOKUP( ${col_question}{row}&" : Label", \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_index_column_range}, \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_values_range} )'.format(**substitutes)
                # category_analysisvalue_lookup_validation_array = '_xlfn.XLOOKUP( ${col_question}{row}&" : Validation", \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_index_column_range}, \'{sheet_name_analysisvalues}\'!{sheet_analysisvalues_values_range} )'.format(**substitutes)
                category_analysisvalue_lookup_row_index = '_xlfn.MATCH( ${col_category}{row}, {category_analysisvalue_lookup_categorynames_array}, 0 )'.format(**substitutes,category_analysisvalue_lookup_categorynames_array=category_analysisvalue_lookup_categorynames_array)

                category_label_mdd = mdmcategory.Label
                try:
                    category_label_prev = prev_map.read_category_label(category_question_name,category_name)
                except prev_map.CellNotFound:
                    category_label_prev = None
                category_label_lookup = '_xlfn.INDEX({where_looking_at}, {index_looking_for} )'.format(where_looking_at=category_analysisvalue_lookup_labels_array,index_looking_for=category_analysisvalue_lookup_row_index)
                category_label = '=IF({val}&""="","",{val}&"")'.format(val=category_label_lookup)
                # category_label = ArrayFormula(text=category_label,ref='')

                category_analysisvalue_mdd = aa_logic.sanitize_analysis_value(mdmcategory.Properties['Value'])
                try:
                    category_analysisvalue_prev = prev_map.read_category_analysisvalue(category_question_name,category_name)
                except prev_map.CellNotFound:
                    category_analysisvalue_prev = None
                category_analysisvalue_lookup = '_xlfn.INDEX({where_looking_at}, {index_looking_for} )'.format(where_looking_at=category_analysisvalue_lookup_analysisvalues_array,index_looking_for=category_analysisvalue_lookup_row_index)
                category_analysisvalue = '=IF({val}&""="","",{val}&"")'.format(val=category_analysisvalue_lookup)
                # category_analysisvalue = ArrayFormula(text=category_analysisvalue,ref='')

                category_validation = '=IF({col_helper_validation_varnotinexclusions}{row},IF(ISERROR(${col_analysisvalue}{row}),FALSE,AND({col_helper_validation_notblank}{row},{col_helper_validation_iswhole}{row},{col_helper_validation_isunique}{row},OR(ISBLANK(${col_sharedlist}{row}),{col_helper_validation_isuniquewithinsl}{row}))),TRUE)'.format(**substitutes)
                
                category_helper_validation_varnotinexclusions = '=VLOOKUP(${col_question}{row},\'{sheet_name_mdddata_variables}\'!$A$2:${col_mdddata_variables_include}$999999,{col_mdddata_variables_include_vlookup_index},FALSE)'.format(**substitutes)
                category_helper_validation_notblank           = '=IF(NOT(ISERROR(${col_analysisvalue}{row})),NOT(OR(ISBLANK({col_analysisvalue}{row}),{col_analysisvalue}{row}="")),FALSE)'.format(**substitutes)
                category_helper_validation_iswhole            = '=IF(ISERROR(VALUE({col_analysisvalue}{row})),FALSE,AND(VALUE({col_analysisvalue}{row})>=0,MOD(VALUE({col_analysisvalue}{row}),1)=0))'.format(**substitutes)
                category_helper_validation_isunique           = '=MATCH(${col_helper_field_categoryvalue}{row},${col_helper_field_categoryvalue}$2:${col_helper_field_categoryvalue}$999999,0)=MATCH(${col_helper_field_categorypath}{row},${col_helper_field_categorypath}$2:${col_helper_field_categorypath}$999999,0)'.format(**substitutes)
                category_helper_validation_isuniquewithinsl   = '=IF(ISBLANK(${col_sharedlist}{row}),TRUE,MATCH(${col_helper_field_slvalue}{row},${col_helper_field_slvalue}$2:${col_helper_field_slvalue}$999999,0)=MATCH(${col_helper_field_slpath}{row},${col_helper_field_slpath}$2:${col_helper_field_slpath}$999999,0))'.format(**substitutes)

                category_helper_field_categorypath    = '=IF(NOT(ISBLANK(${col_analysisvalue}{row})),""&""&${col_question}{row}&".Categories["&${col_category}{row}&"]","")'.format(**substitutes)
                category_helper_field_categoryvalue   = '=IF(NOT(ISBLANK(${col_analysisvalue}{row})),""&""&${col_question}{row}&".AnalysisValues["&${col_analysisvalue}{row}&"]","")'.format(**substitutes)
                category_helper_field_slpath          = '=IF(NOT(ISBLANK(${col_sharedlist}{row})),""&"SharedList["&${col_sharedlist}{row}&"].Categories["&${col_category}{row}&"]","-")'.format(**substitutes)
                category_helper_field_slvalue         = '=IF(NOT(ISBLANK(${col_analysisvalue}{row})),IF(NOT(ISBLANK(${col_sharedlist}{row})),""&"SharedList["&${col_sharedlist}{row}&"].AnalysisValues["&${col_analysisvalue}{row}&"]","-"),"")'.format(**substitutes)

                row_add = [

                    category_path,

                    category_question_name,
                    category_name,

                    category_sharedlist,

                    category_label_mdd,
                    category_label_prev,
                    category_label,

                    category_analysisvalue_mdd,
                    category_analysisvalue_prev,
                    category_analysisvalue,

                    category_validation,

                    category_helper_validation_varnotinexclusions,
                    category_helper_validation_notblank,
                    category_helper_validation_iswhole,
                    category_helper_validation_isunique,
                    category_helper_validation_isuniquewithinsl,
                    category_helper_field_categorypath,
                    category_helper_field_categoryvalue,
                    category_helper_field_slpath,
                    category_helper_field_slvalue,
                ]
                data.append(*row_add)

    return data.to_df()




