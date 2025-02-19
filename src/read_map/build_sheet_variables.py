

import re



if __name__ == '__main__':
    # run as a program
    import util_performance_monitor as util_perf
    import util_mdmvars
    import util_dataframe_wrapper
    import columns_sheet_variables as sheet
    import columns_sheet_mdddata_variables
elif '.' in __name__:
    # package
    from . import util_performance_monitor as util_perf
    from . import util_mdmvars
    from . import util_dataframe_wrapper
    from . import columns_sheet_variables as sheet
    from . import columns_sheet_mdddata_variables
else:
    # included with no parent package
    import util_performance_monitor as util_perf
    import util_mdmvars
    import util_dataframe_wrapper
    import columns_sheet_variables as sheet
    import columns_sheet_mdddata_variables



def build_df(mdmvariables,df_prev,config):
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
        if not util_mdmvars.is_data_item(mdmvariable):
            continue
        if util_mdmvars.is_nocasedata(mdmvariable):
            continue
    
        substitutes = {
            **sheet.column_letters,
            'row': data.get_working_row_number(),
            'sheet_name_mdddata': columns_sheet_mdddata_variables.sheet_name,
            'col_mdddata_question_type': '',
            # 'col_mdddata_col_last': columns_sheet_mdddata_variables.column_letters['col_validation'],
            'col_mdddata_question_type': columns_sheet_mdddata_variables.column_letters['col_question_type'],
            'col_mdddata_question_type_vlookup_index': columns_sheet_mdddata_variables.column_vlookup_index['col_question_type'],
            'col_mdddata_validation': columns_sheet_mdddata_variables.column_letters['col_validation'],
            'col_mdddata_validation_vlookup_index': columns_sheet_mdddata_variables.column_vlookup_index['col_validation'],
        }

        question_name = mdmvariable.FullName

        question_type = '=VLOOKUP($A{row},\'{sheet_name_mdddata}\'!$A$2:${col_mdddata_question_type}$999999,{col_mdddata_question_type_vlookup_index},FALSE)&""'.format(**substitutes)

        question_shortname = \
            ( \
                df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_shortname_prev']] \
                if df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_shortname_prev']] \
                else df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_shortname_mdd']] \
            ) \
            if df_prev is not None and question_name in df_prev.index.get_level_values(0) \
            else ''

        question_label = \
            ( \
                df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_label_prev']] \
                if df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_label_prev']] \
                else df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_label_mdd']] \
            ) \
            if df_prev is not None and question_name in df_prev.index.get_level_values(0) \
            else ''

        question_include_truefalse = \
            ( \
                df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_include_prev']] \
                if df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_include_prev']] \
                else df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_include_mdd']] \
            ) \
            if df_prev is not None and question_name in df_prev.index.get_level_values(0) \
            else ''
        
        question_include = \
            '' if question_include_truefalse=='' \
            else ( 'x' if question_include_truefalse==True or re.match(r'^\s*?(?:true)\s*?$','{s}'.format(s=question_include_truefalse),flags=re.I|re.DOTALL) \
            else ( '' if question_include_truefalse==False or re.match(r'^\s*?(?:false)\s*?$','{s}'.format(s=question_include_truefalse),flags=re.I|re.DOTALL) else '?' ) )

        question_validation_truefalse = 'VLOOKUP($A{row},\'{sheet_name_mdddata}\'!$A$2:${col_mdddata_validation}$999999,{col_mdddata_validation_vlookup_index},FALSE)'.format(**substitutes)
        question_validation = '=IF({val},"","Failed")'.format(val=question_validation_truefalse)

        question_comment = \
            df_prev.loc[question_name,columns_sheet_mdddata_variables.column_names['col_comment_prev']] \
            if df_prev is not None and question_name in df_prev.index.get_level_values(0) \
            else ''

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




