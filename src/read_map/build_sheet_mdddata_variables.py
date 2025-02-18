
import re



if __name__ == '__main__':
    # run as a program
    import aa_logic
    import util_performance_monitor as util_perf
    import util_mdmvars
    import util_dataframe_wrapper
    import columns_sheet_variables
    import columns_sheet_mdddata_variables as sheet
    import columns_sheet_mdddata_categories
elif '.' in __name__:
    # package
    from . import aa_logic
    from . import util_performance_monitor as util_perf
    from . import util_mdmvars
    from . import util_dataframe_wrapper
    from . import columns_sheet_variables
    from . import columns_sheet_mdddata_variables as sheet
    from . import columns_sheet_mdddata_categories
else:
    # included with no parent package
    import aa_logic
    import util_performance_monitor as util_perf
    import util_mdmvars
    import util_dataframe_wrapper
    import columns_sheet_variables
    import columns_sheet_mdddata_variables as sheet
    import columns_sheet_mdddata_categories



def parse_shortname_formula(txt):
    def repl_cat_analysisvalue(matches):
        return '"&VLOOKUP("{var}.Categories[{cat}]",\'{sheet_name}\'!$A2:${col_last}999999,{vlookup_index},FALSE)&"'.format(var=matches[1],cat=matches[2],sheet_name=columns_sheet_mdddata_categories.sheet_name,col_last=columns_sheet_mdddata_categories.column_letters['col_analysisvalue'],vlookup_index=columns_sheet_mdddata_categories.column_vlookup_index['col_analysisvalue'])
    def repl_var_shortname(matches):
        return '"&VLOOKUP("{var}",\'{sheet_name}\'!$A2:${col_last}999999,{vlookup_index},FALSE)&"'.format(var=matches[1],sheet_name=sheet.sheet_name,col_last=sheet.column_letters['col_shortname'],vlookup_index=sheet.column_vlookup_index['col_shortname'])
    if '[L' in txt:
        assert '"' not in txt, '\'"\' in shortname, please check'
        return '="{f}"'.format(f=re.sub(r'\[L:(.*?)\]',repl_var_shortname,re.sub(r'\[L:(.*?):(.*?)\]',repl_cat_analysisvalue,txt,flags=re.I|re.DOTALL),flags=re.I|re.DOTALL))
    else:
        return txt



def build_df(mdmvariables,df_prev,config):
    data = util_dataframe_wrapper.PandasDataframeWrapper(sheet.columns)

    performance_counter = iter(util_perf.PerformanceMonitor(config={
        'total_records': len(mdmvariables),
        'report_frequency_records_count': 1,
        'report_frequency_timeinterval': 14,
        'report_text_pipein': 'progress reading derived ShortName for variables',
    }))
    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue
        next(performance_counter)

        substitutes = {
            **sheet.column_letters,
            'row': data.get_working_row_number(),
            'sheet_name_variables': columns_sheet_variables.sheet_name,
            'col_variablessheet_col_last': columns_sheet_variables.column_letters[columns_sheet_variables.column_aliases[-1]],
            'col_variablessheet_label_vlookup_index': columns_sheet_variables.column_vlookup_index['col_label'],
            'col_variablessheet_shortname_vlookup_index': columns_sheet_variables.column_vlookup_index['col_shortname'],
            'col_variablessheet_include_vlookup_index': columns_sheet_variables.column_vlookup_index['col_include'],
            'col_variablessheet_comment_vlookup_index': columns_sheet_variables.column_vlookup_index['col_comment'],
        }

        question_name = mdmvariable.FullName

        question_data_items_within_parent_wing = \
            [ '$A{row}'.format(**substitutes) ] \
            if util_mdmvars.is_data_item(mdmvariable) \
            else [
                '"{name}"'.format(name=mdmitem.FullName) for mdmitem in util_mdmvars.list_mdmdatafields_recursively(mdmvariable)
            ]

        question_type = util_mdmvars.compile_question_type_description(mdmvariable)

        question_isdatavariable = True if util_mdmvars.is_data_item(mdmvariable) else False

        question_label_mdd = '{s}'.format(s=mdmvariable.Label)

        question_label_prev = df_prev.loc[question_name,sheet.column_names['col_label']] if df_prev is not None and question_name in df_prev.index.get_level_values(0) else ''

        question_label_lookup = 'VLOOKUP($A{row},\'{sheet_name_variables}\'!$A$2:${col_variablessheet_col_last}$999999,{col_variablessheet_label_vlookup_index},FALSE)&""'.format(**substitutes)
        question_label = '=IF(ISERROR({val}),${col_label_mdd}{row}&"",{val})'.format(**substitutes,val=question_label_lookup)

        question_shortname_mdd = parse_shortname_formula(aa_logic.read_shortname(mdmvariable))

        question_shortname_prev = parse_shortname_formula( df_prev.loc[question_name,sheet.column_names['col_shortname']] if df_prev is not None and question_name in df_prev.index.get_level_values(0) else '' )

        question_shortname = '=' + '&", "&'.join([
            'VLOOKUP({item_name},\'{sheet_name_variables}\'!$A$2:${col_variablessheet_col_last}$999999,{col_variablessheet_shortname_vlookup_index},FALSE)&""'.format(**substitutes,item_name='$A{row}'.format(**substitutes) if item_name==question_name else item_name) for item_name in question_data_items_within_parent_wing
        ]) + '&""'
        question_shortname = question_shortname.replace('=&','=""&')

        question_comment_prev = df_prev.loc[question_name,sheet.column_names['col_comment']] if df_prev is not None and question_name in df_prev.index.get_level_values(0) else ''

        question_comment_lookup = 'VLOOKUP($A{row},\'{sheet_name_variables}\'!$A$2:${col_variablessheet_col_last}$999999,{col_variablessheet_comment_vlookup_index},FALSE)'.format(**substitutes)
        question_comment = '=IF(ISERROR({val}),"",{val})'.format(val=question_comment_lookup)

        question_include_mdd = '' if not util_mdmvars.is_data_item(mdmvariable) else ( False if aa_logic.should_exclude(mdmvariable) else True )

        question_include_prev = df_prev.loc[question_name,sheet.column_names['col_include']] if df_prev is not None and question_name in df_prev.index.get_level_values(0) else ''

        question_include_lookup = 'IF(' + '&'.join(['""']+[
            'VLOOKUP({item_name},\'{sheet_name_variables}\'!$A$2:${col_variablessheet_col_last}$999999,{col_variablessheet_include_vlookup_index},FALSE)&""'.format(**substitutes,item_name='$A{row}'.format(**substitutes) if item_name==question_name else item_name) for item_name in question_data_items_within_parent_wing
        ]) + '&""="",FALSE,TRUE)'
        question_include = '=IF(ISERROR({val}),FALSE,{val})'.format(val=question_include_lookup)

        # question_validation_field_shortnamelcase = '=IF($F{row},LOWER($E{row}),"")'.format(row=row)
        # question_validation_notblank = '=IF($F{row},AND(NOT(ISBLANK($E{row})),NOT($E{row}=0)),"")'.format(row=row)
        # question_validation_isalphanumeric= '=IF($F{row},LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(${col_helper_shortname_clean_lcase}{row},"a",""),"b",""),"c",""),"d",""),"e",""),"f",""),"g",""),"h",""),"i",""),"j",""),"k",""),"l",""),"m",""),"n",""),"o",""),"p",""),"q",""),"r",""),"s",""),"t",""),"u",""),"v",""),"w",""),"x",""),"y",""),"z",""),"0",""),"1",""),"2",""),"3",""),"4",""),"5",""),"6",""),"7",""),"8",""),"9",""),"_",""))=0,"")'.format(row=row)
        # question_validation_notstartszero = '=IF($F{row},NOT(OR(LEFT(${col_helper_shortname_clean_lcase}{row},"1")="0",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="1",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="2",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="3",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="4",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="5",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="6",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="7",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="8",LEFT(${col_helper_shortname_clean_lcase}{row},"1")="9")),"")'.format(row=row)
        # question_validation_isunique = '=IF($F{row},(MATCH($A{row},$A$2:$A$999999,0)=MATCH(${col_helper_shortname_clean_lcase}{row},$I$2:$I$999999,0)),"")'.format(row=row)
        question_validation = ('=IF(' \
            + 'NOT(${col_question_isdatavariable}{row}),' \
            + 'TRUE,' \
            + 'IF(' \
                + 'NOT(${col_include}{row}),' \
                + 'TRUE,' \
                + 'IF(' \
                    + 'ISERROR(${col_shortname}{row}),' \
                    + 'FALSE,' \
                    + 'AND(' \
                        + '${col_helper_validation_notblank}{row},' \
                        + '${col_helper_validation_isalphanumeric}{row},' \
                        + '${col_helper_validation_notstartswithnumber}{row},' \
                        + '${col_helper_validation_lengthcheck}{row},' \
                        + '${col_helper_validation_reservedkeywordcheck}{row},' \
                        + '${col_helper_validation_isunique}{row}' \
                    + ')' \
                + ')' \
            + ')' \
        + ')').format(**substitutes)
        question_helper_shortname_clean_lcase = '=IF(AND(${col_question_isdatavariable}{row},${col_include}{row}),LOWER(${col_shortname}{row}),"-")'.format(**substitutes)
        question_helper_validation_notblank = '=AND(NOT(ISBLANK(${col_helper_shortname_clean_lcase}{row})),NOT(${col_helper_shortname_clean_lcase}{row}=""))'.format(**substitutes)
        question_helper_validation_isalphanumeric = '=(LEN(<<VAR>>)=0)'.format(**substitutes)
        for letter in ['a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','0','1','2','3','4','5','6','7','8','9','_','$','#','@','.']:
            question_helper_validation_isalphanumeric = question_helper_validation_isalphanumeric.replace('<<VAR>>','SUBSTITUTE(<<VAR>>,"{letter}","")'.format(letter=letter))
        question_helper_validation_isalphanumeric = question_helper_validation_isalphanumeric.replace('<<VAR>>','${col_helper_shortname_clean_lcase}{row}'.format(**substitutes))
        question_helper_validation_notstartswithnumber = ('=NOT(OR(RIGHT(${col_helper_shortname_clean_lcase}{row},1)=".",'+','.join([ 'LEFT(${col_helper_shortname_clean_lcase}{row},1)="{letter}"'.format(**substitutes,letter=letter) for letter in ['0','1','2','3','4','5','6','7','8','9','$','#','@','.'] ])+'))').format(**substitutes)
        question_helper_validation_lengthcheck = '=LEN(${col_helper_shortname_clean_lcase}{row})<=32'.format(**substitutes)
        question_helper_validation_reservedkeywordcheck = ('=NOT(OR('+','.join(['${col_helper_shortname_clean_lcase}{row}="{word}"'.format(**substitutes,word=word) for word in ['ALL','AND','BY','EQ','GE','GT','LE','LT','NE','NOT','OR','TO','WITH',] ])+'))').format(**substitutes)
        question_helper_validation_isunique = '=(MATCH($A{row},$A$2:$A$999999,0)=MATCH(${col_helper_shortname_clean_lcase}{row},${col_helper_shortname_clean_lcase}$2:${col_helper_shortname_clean_lcase}$999999,0))'.format(**substitutes)

        row_add = [
            question_name,
            question_type,
            question_isdatavariable,
            question_label_mdd,
            question_label_prev,
            question_label,
            question_shortname_mdd,
            question_shortname_prev,
            question_shortname,
            question_comment_prev,
            question_comment,
            question_include_mdd,
            question_include_prev,
            question_include,
            question_validation,
            question_helper_shortname_clean_lcase,
            question_helper_validation_notblank,
            question_helper_validation_isalphanumeric,
            question_helper_validation_notstartswithnumber,
            question_helper_validation_lengthcheck,
            question_helper_validation_reservedkeywordcheck,
            question_helper_validation_isunique,
        ]
        data.append(*row_add)

    return data.to_df()



