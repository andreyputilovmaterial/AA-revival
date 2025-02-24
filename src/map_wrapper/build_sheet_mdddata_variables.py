
import re



from . import aa_logic
from . import util_performance_monitor as util_perf
from . import util_dataframe_wrapper
from .column_definitions import sheet_variables
from .column_definitions import sheet_mdddata_variables as sheet
from .column_definitions import sheet_mdddata_categories




CONFIG_VALIDATION_MAX_SHORTNAME_LENGTH = 50




def helper_parse_shortname_formula(txt):
    def repl_cat_analysisvalue(matches):
        substitute_part = 'VLOOKUP("{var}.Categories[{cat}]",\'{sheet_name}\'!$A2:${col_last}999999,{vlookup_index},FALSE)'.format(var=matches[1],cat=matches[2],sheet_name=sheet_mdddata_categories.sheet_name,col_last=sheet_mdddata_categories.column_letters['col_analysisvalue'],vlookup_index=sheet_mdddata_categories.column_vlookup_index['col_analysisvalue'])
        substitute_part = 'IF(LEN({A})=1,"00"&{A},IF(LEN({A})=2,"0"&{A},{A}))'.format(A=substitute_part)
        return '"&{formula}&"'.format(formula=substitute_part)
    def repl_var_shortname(matches):
        substitute_part = 'VLOOKUP("{var}",\'{sheet_name}\'!$A2:${col_last}999999,{vlookup_index},FALSE)'.format(var=matches[1],sheet_name=sheet.sheet_name,col_last=sheet.column_letters['col_shortname'],vlookup_index=sheet.column_vlookup_index['col_shortname'])
        substitute_part = 'IF(LEN({A})=1,"00"&{A},IF(LEN({A})=2,"0"&{A},{A}))'.format(A=substitute_part)
        return '"&{formula}&"'.format(formula=substitute_part)
    if '[L:' in txt:
        assert '"' not in txt, '\'"\' in shortname, please check'
        return '="{f}"'.format(f=re.sub(r'\[L:(.*?)\]',repl_var_shortname,re.sub(r'\[L:(.*?):(.*?)\]',repl_cat_analysisvalue,txt,flags=re.I|re.DOTALL),flags=re.I|re.DOTALL))
    else:
        return txt



def build_df(mdd,prev_map,config):
    mdmvariables = mdd.variables
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
            'sheet_name_variables': sheet_variables.sheet_name,
            'col_variablessheet_col_last': sheet_variables.column_letters[sheet_variables.column_aliases[-1]],
            'col_variablessheet_label_vlookup_index': sheet_variables.column_vlookup_index['col_label'],
            'col_variablessheet_shortname_vlookup_index': sheet_variables.column_vlookup_index['col_shortname'],
            'col_variablessheet_include_vlookup_index': sheet_variables.column_vlookup_index['col_include'],
            'col_variablessheet_comment_vlookup_index': sheet_variables.column_vlookup_index['col_comment'],
        }

        question_name = mdmvariable.FullName

        question_data_items_within_parent_wing = \
            [ '$A{row}'.format(**substitutes) ] \
            if mdd.is_data_item(mdmvariable) \
            else [
                '"{name}"'.format(name=mdmitem.FullName) for mdmitem in mdd.list_mdmdatafields_recursively(mdmvariable)
            ]

        question_type = mdd.compile_question_type_description(mdmvariable)

        question_isdatavariable = True if mdd.is_data_item(mdmvariable) else False

        question_label_mdd = '{s}'.format(s=mdmvariable.Label)

        try:
            question_label_prev = prev_map.read_question_label(question_name)
        except prev_map.CellNotFound:
            question_label_prev = None

        question_label_lookup = 'VLOOKUP($A{row},\'{sheet_name_variables}\'!$A$2:${col_variablessheet_col_last}$999999,{col_variablessheet_label_vlookup_index},FALSE)&""'.format(**substitutes)
        question_label = '=IF(ISERROR({val}),${col_label_mdd}{row}&"",{val})'.format(**substitutes,val=question_label_lookup)

        question_shortname_mdd = helper_parse_shortname_formula(aa_logic.read_shortname(mdmvariable,mdd))

        try:
            question_shortname_prev = helper_parse_shortname_formula( prev_map.read_question_shortname(question_name) )
        except prev_map.CellNotFound:
            question_shortname_prev = None

        question_shortname = '=' + '&", "&'.join([
            'VLOOKUP({item_name},\'{sheet_name_variables}\'!$A$2:${col_variablessheet_col_last}$999999,{col_variablessheet_shortname_vlookup_index},FALSE)&""'.format(**substitutes,item_name='$A{row}'.format(**substitutes) if item_name==question_name else item_name) for item_name in question_data_items_within_parent_wing
        ]) + '&""'
        question_shortname = question_shortname.replace('=&','=""&')

        try:
            question_comment_prev = prev_map.read_question_comment(question_name)
        except prev_map.CellNotFound:
            question_comment_prev = None

        question_comment_lookup = 'VLOOKUP($A{row},\'{sheet_name_variables}\'!$A$2:${col_variablessheet_col_last}$999999,{col_variablessheet_comment_vlookup_index},FALSE)&""'.format(**substitutes)
        question_comment = '=IF(ISERROR({val}),"",{val})'.format(val=question_comment_lookup)

        question_include_mdd = '' if not mdd.is_data_item(mdmvariable) else ( False if aa_logic.should_exclude(mdmvariable,mdd) else True )

        try:
            question_include_prev = prev_map.read_question_is_included(question_name)
        except prev_map.CellNotFound:
            question_include_prev = None

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
        question_helper_validation_lengthcheck = '=LEN(${col_helper_shortname_clean_lcase}{row})<={MAXLENGTH}'.format(**substitutes,MAXLENGTH=CONFIG_VALIDATION_MAX_SHORTNAME_LENGTH)
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



