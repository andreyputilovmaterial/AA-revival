
import re
from pathlib import Path
import pandas as pd
try:
    from zoneinfo import ZoneInfo
    import time
    get_localzone_name = None
    try:
        from tzlocal import get_localzone_name
    except:
        pass # if it's not installed, I will not require it, it's not essential
except:
    ZoneInfo = None
    pass # older python
from openpyxl.utils import get_column_letter
# from openpyxl.worksheet.formula import ArrayFormula


if __name__ == '__main__':
    # run as a program
    import util_performance_monitor as util_perf
    import mdmutils
    import aa_logic
elif '.' in __name__:
    # package
    from . import util_performance_monitor as util_perf
    from . import mdmutils
    from . import aa_logic
else:
    # included with no parent package
    import util_performance_monitor as util_perf
    import mdmutils
    import aa_logic





CONFIG_TOOL_NAME = 'IPS Tools - Author Specifications Revival v0.000'




def build_overview_df(variables,df,config):
    # if df:
    #     df.drop(df.index,inplace=True)
    df = pd.DataFrame(
        data = {
            'Value': [],
        },
        index = pd.Series( [], name = 'Description' ),
    )
    mdd_filename_part = Path(config['mdd_filename']).stem
    mdd_refresh_datetime = config['datetime']
    try:
        mdd_refresh_datetime = mdd_refresh_datetime.strftime('%m/%d/%Y %I:%M:%S %p %Z')
        if ZoneInfo:
            try:
                mdd_refresh_datetime_naive = config['datetime']
                try:
                    system_tz = ZoneInfo(time.tzname[0])
                except Exception as e:
                    if get_localzone_name:
                        system_tz = ZoneInfo(get_localzone_name())
                    else:
                        raise e
                mdd_refresh_datetime_local = mdd_refresh_datetime_naive.replace(tzinfo=system_tz)
                mdd_refresh_datetime_local = mdd_refresh_datetime_local.strftime('%m/%d/%Y %I:%M:%S %p %Z')
                mdd_refresh_datetime = mdd_refresh_datetime_local
            except:
                pass
    except:
        pass
    mdmdoc = variables[0]
    df.loc['Tool'] = [CONFIG_TOOL_NAME]
    df.loc['MDD'] = [mdd_filename_part]
    df.loc['Validation'] = ['=IF(AND(AND(Variables!$G$2:$G$699999),AND(MDD_Data_Categories!$G$2:$G$699999)),"No issues","Failed")']
    df.loc['MDD path'] = [config['mdd_filename']]
    df.loc['MDD last refreshed'] = [mdd_refresh_datetime]
    df.loc['JobNumber property from MDD'] = [mdmdoc.Properties['JobNumber']]
    df.loc['StudyName propertty from MDD'] = [mdmdoc.Properties['StudyName']]
    df.loc['Client property from MDD'] = [mdmdoc.Properties['Client']]
    return df



# this fn generates human readable description of what variable type is
# the expected result could be just "text" or "single-punch"
# or "single-punch grid" or "numeric grid"
# or "datetime, inside a loop, inside a block with fields..."
# or "text, a helper field of..."
# so it the logic sometimes looks complicated, it calls it recursively and conditionally depending on parent
# but this function is in fact not important, it only generates that text that is not used programmatically
# so if it's something too complex to understand and support, you can just replace this whole function with some more straightforward implementation
# so these complex codes can be safely deleted (I mean, all within compile_question_type_description())
def compile_question_type_description(mdmvar):
    def compile_type_description_this_level_field(mdmvar):
        if mdmutils.has_parent(mdmvar):
            count_sisters = 0
            mdmparent = mdmutils.get_parent(mdmvar)
            is_helperfield = mdmparent.ObjectTypeValue==0
            if not is_helperfield:
                for f in mdmparent.Fields:
                    if mdmutils.is_dataless(f):
                        continue
                    elif f.Name==mdmvar.Name:
                        continue
                    count_sisters = count_sisters + 1
            if mdmutils.is_iterative(mdmparent) and count_sisters == 0:
                if mdmutils.is_data_item(mdmvar):
                    if mdmvar.DataType==3 and mdmvar.MaxValue==1:
                        return 'single-punch grid'
                    elif mdmvar.DataType==3:
                        return 'multi-punch grid'
                    elif mdmvar.DataType==1 or mdmvar.DataType==6:
                        return 'numeric grid'
                    elif mdmvar.DataType==2:
                        return 'text grid'
                    elif mdmvar.DataType==5:
                        return 'datetime grid'
                    elif mdmvar.DataType==7:
                        return 'boolean grid'
        if mdmvar.ObjectTypeValue==27:
            return 'mdm document'
        elif mdmvar.ObjectTypeValue==1 or mdmvar.ObjectTypeValue==2:
            return 'loop'
        elif mdmvar.ObjectTypeValue==3:
            return 'block with fields'
        elif mdmvar.ObjectTypeValue==0:
            if mdmvar.DataType==0:
                return 'not a data variable (info node)'
            elif mdmvar.DataType==1 or mdmvar.DataType==6:
                return 'numeric'
            elif mdmvar.DataType==2:
                return 'text'
            elif mdmvar.DataType==3 and mdmvar.MaxValue==1:
                return 'single-punch'
            elif mdmvar.DataType==3:
                return 'multi-punch'
            elif mdmvar.DataType==4:
                return 'object'
            elif mdmvar.DataType==5:
                return 'datetime'
            elif mdmvar.DataType==7:
                return 'boolean flag'
            elif mdmvar.DataType==16:
                return 'ivariables instance'
            else:
                raise Exception('unrecognized DataType: {o}'.format(o=mdmvar.DataType))
        raise Exception('unrecognized ObjectTypeValue: {o}'.format(o=mdmvar.ObjectTypeValue))
    result = compile_type_description_this_level_field(mdmvar)
    mdmparent = mdmutils.get_parent(mdmvar) if mdmutils.has_parent(mdmvar) else None
    parent_part_processed = False
    if re.match(r'.*?\bgrid\b.*?',result):
        parent_part_processed = True
    if mdmutils.is_helper_field(mdmvar):
        parent_part_processed
        result = result + ', a helper field of var of type '+compile_question_type_description(mdmparent)
        return result
    while mdmparent:
        if not parent_part_processed:
            parent_part = compile_type_description_this_level_field(mdmparent)
            result = result + ', inside {article}{type}'.format(type=parent_part,article='a ' if re.match(r'^\s*?(?:loop|array|block).*?',parent_part) else '')
        mdmparent = mdmutils.get_parent(mdmparent) if mdmutils.has_parent(mdmparent) else None
        parent_part_processed = False
    return result

def build_mdd_data_variables_df(mdmvariables,df_prev,config):
    def parse_shortname_formula(txt):
        def repl_cat_analysisvalue(matches):
            return '"&VLOOKUP("{var}.Categories[{cat}]",\'MDD_Data_Categories\'!$A2:$G699999,6,FALSE)&"'.format(var=matches[1],cat=matches[2])
        def repl_var_shortname(matches):
            return '"&VLOOKUP("{var}",\'MDD_Data_Variables\'!$A2:$G699999,5,FALSE)&"'.format(var=matches[1])
        if '[L' in txt:
            assert '"' not in txt, '\'"\' in shortname, please check'
            return '="{f}"'.format(f=re.sub(r'\[L:(.*?)\]',repl_var_shortname,re.sub(r'\[L:(.*?):(.*?)\]',repl_cat_analysisvalue,txt,flags=re.I|re.DOTALL),flags=re.I|re.DOTALL))
        else:
            return txt
    df = pd.DataFrame(
        data = {
            'Question': [],
            'Question Type': [],
            'ShortName in MDD': [],
            'Label': [],
            'Suggested ShortName': [],
            'Helper Validation Field - Is data variable' : [],
            'Validation': [],
            'Helper Validation ShortName Lowercase Field' : [],
            'Helper Validation - Not blank' : [],
            'Helper Validation - Is alphanumeric' : [],
            'Helper Validation - Does not start with a number' : [],
            'Helper Validation - Is unique' : [],
        },
        index = None,
    )
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
        'report_text_pipein': 'progress reading derived ShortName for variables',
    }))
    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue
        question_name = mdmvariable.FullName
        next(performance_counter)
        question_shortname = None
        if mdmutils.is_data_item(mdmvariable):
            try:
                if mdmvariable.FullName in df_prev.index.get_level_values(0):
                    question_shortname = df_prev.loc[mdmvariable.FullName,'Suggested ShortName']
            except Exception as e:
                raise e
                # pass
            if not question_shortname:
                question_shortname = aa_logic.read_shortname(mdmvariable)
        if not question_shortname:
            question_shortname = ''
        question_shortname = parse_shortname_formula(question_shortname)
        question_type = compile_question_type_description(mdmvariable)
        question_label = '{s}'.format(s=mdmvariable.Label)

        # question_suggested_shortname = '=VLOOKUP($A{row},Variables!$A$2:$H$699999,4,FALSE)'.format(row=row) if mdmutils.is_data_item(mdmvariable) else ''
        question_suggested_shortname = ''
        if mdmutils.is_data_item(mdmvariable):
            # a data-item - show ShortName with VLOOKUP()
            question_suggested_shortname = '=VLOOKUP($A{row},Variables!$A$2:$H$699999,4,FALSE)'.format(row=row)
        else:
            # not a data items - compile a comma-separated list of all data-items inside
            question_suggested_shortname = '=""'
            for mdm_item_within in mdmutils.list_mdmdatafields_recursively(mdmvariable):
                question_suggested_shortname = question_suggested_shortname+ '&", "&VLOOKUP("{var}",Variables!$A$2:$H$699999,4,FALSE)'.format(row=row,var=mdm_item_within.FullName)
            question_suggested_shortname = question_suggested_shortname.replace('=""&", "','=""')

        question_validation_field_isdatavariable = 'TRUE' if mdmutils.is_data_item(mdmvariable) else 'FALSE'

        question_validation = '=IF($F{row},IF(NOT(ISERROR($E{row})),AND(I{row},J{row},K{row},L{row}),FALSE),"")'.format(row=row) if mdmutils.is_data_item(mdmvariable) else ''

        question_validation_field_shortnamelcase = '=IF($F{row},LOWER($E{row}),"")'.format(row=row)
        question_validation_notblank = '=IF($F{row},AND(NOT(ISBLANK($E{row})),NOT($E{row}=0)),"")'.format(row=row)
        question_validation_isalphanumeric= '=IF($F{row},LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{row},"a",""),"b",""),"c",""),"d",""),"e",""),"f",""),"g",""),"h",""),"i",""),"j",""),"k",""),"l",""),"m",""),"n",""),"o",""),"p",""),"q",""),"r",""),"s",""),"t",""),"u",""),"v",""),"w",""),"x",""),"y",""),"z",""),"0",""),"1",""),"2",""),"3",""),"4",""),"5",""),"6",""),"7",""),"8",""),"9",""),"_",""))=0,"")'.format(row=row)
        question_validation_notstartszero = '=IF($F{row},NOT(OR(LEFT(H{row},"1")="0",LEFT(H{row},"1")="1",LEFT(H{row},"1")="2",LEFT(H{row},"1")="3",LEFT(H{row},"1")="4",LEFT(H{row},"1")="5",LEFT(H{row},"1")="6",LEFT(H{row},"1")="7",LEFT(H{row},"1")="8",LEFT(H{row},"1")="9")),"")'.format(row=row)
        question_validation_isunique = '=IF($F{row},(MATCH($A{row},$A$2:$A$699999,0)=MATCH($H{row},$H$2:$H$699999,0)),"")'.format(row=row)

        row_add_list = [
            
            question_name,

            question_type,
            question_shortname,
            question_label,
            question_suggested_shortname,
            question_validation_field_isdatavariable,

            question_validation,
            
            question_validation_field_shortnamelcase,
            question_validation_notblank,
            question_validation_isalphanumeric,
            question_validation_notstartszero,
            question_validation_isunique,
        ]
        row_add = {col[0]: col[1] for col in zip(df.columns,row_add_list)}
        data_add.append(row_add)
        row = row + 1
    
    df = pd.concat([df,pd.DataFrame(data_add)],ignore_index=True)
    df.set_index('Question',inplace=True)
    return df



def build_mdd_data_categories_df(mdmvariables,df,config):
    df = pd.DataFrame(
        data = {
            'Category_Path': [],
            'Variable': [],
            'Category': [],
            'SharedList': [],
            'Label': [],
            'AnalysisValue': [],
            'Validation Final': [],
            'Validation - Not blank': [],
            'Validation - Is whole number': [],
            'Validation - Is unique': [],
            'Variable is not in exclusions': [],
            'Validation - Is unique within shared list': [],
            'Helper field - vaiable name + category name': [],
            'Helper field - vaiable name + analysis value': [],
            'Helper Field - Shared list name + analysis value': [],
            'Helper Field - Shared list name + category name': [],
        },
        index = None,
    )

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
        'report_text_pipein': 'progress reading analysis values and populating categories sheet',
    }))
    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue
        category_question_name = mdmvariable.FullName
        next(performance_counter)
        if mdmutils.is_iterative(mdmvariable) or mdmutils.has_own_categories(mdmvariable):
            for mdmsharedlist_name, mdmcategory in mdmutils.list_mdmcategories_with_slname(mdmvariable):

                category_name = mdmcategory.Name
                # category_item_name = '{var}.Categories[{cat}]'.format(var=category_question_name,cat=category_name)
                category_item_name = '=""&$B{row}&".Categories["&$C{row}&"]"'.format(row=row)

                category_sharedlist = mdmsharedlist_name if mdmsharedlist_name else ''

                category_label = mdmcategory.Label
                category_analysisvalue = '=IF( _xlfn.INDEX( _xlfn.XLOOKUP( $B{row}&" : Analysis Value", \'Analysis Values\'!$A$2:$A$699999, \'Analysis Values\'!$C$2:$ZZ$699999 ), MATCH( $C{row}, _xlfn.XLOOKUP( $B{row}&"", \'Analysis Values\'!$A$2:$A$699999, \'Analysis Values\'!$C$2:$ZZ$699999 ), 0 ) )&""="", "", VALUE( _xlfn.INDEX( _xlfn.XLOOKUP( $B{row}&" : Analysis Value", \'Analysis Values\'!$A$2:$A$699999, \'Analysis Values\'!$C$2:$ZZ$699999 ), MATCH( $C{row}, _xlfn.XLOOKUP( $B{row}&"", \'Analysis Values\'!$A$2:$A$699999, \'Analysis Values\'!$C$2:$ZZ$699999 ), 0 ) )&"" ) )'.format(row=row)
                # category_analysisvalue = ArrayFormula(text=category_analysisvalue,ref='')
                category_validation = '=IF(K{row},AND(H{row},I{row},J{row},OR(ISBLANK($D{row}),$L{row})),TRUE)'.format(row=row)
                category_validation_field_notblank = '=IF(NOT(ISERROR($F{row})),NOT(OR(ISBLANK($F{row}),$F{row}="")),FALSE)'.format(row=row)
                category_validation_field_iswholenumber = '=IF(ISERROR(VALUE($F{row})),FALSE,AND(VALUE($F{row})>=0,MOD(VALUE($F{row}),1)=0))'.format(row=row)
                category_validation_field_isunique = '=MATCH($M{row},$M$2:$M$699999,0)=MATCH($N{row},$N$2:$N$699999,0)'.format(row=row)
                category_validation_field_varnotinexclusions = '=NOT(ISBLANK(VLOOKUP($B{row},Variables!$A$2:$H$699999,6,FALSE)))'.format(row=row)
                category_validation_field_isuniquewithinsl = '=IF(ISBLANK($D{row}),TRUE,MATCH($P{row},$P$2:$P$699999,0)=MATCH($O{row},$O$2:$O$699999,0))'.format(row=row)
                category_validation_field_helper_varnamepluscatname = '=""&$B{row}&"_"&$C{row}'.format(row=row)
                category_validation_field_helper_varnamepluscatvalue = '=IF(NOT(ISERROR(VALUE($F{row}))),""&$B{row}&"_"&$F{row},""&$B{row}&"_"&"missing")'.format(row=row)
                category_validation_field_helper_slnamepluscatname = '=IF(NOT(ISBLANK($D{row})),IF(NOT(ISERROR(VALUE($F{row}))),""&$D{row}&"_"&$F{row},""&$D{row}&"_"&"missing"),"")'.format(row=row)
                category_validation_field_helper_slnamepluscatvalue = '=IF(NOT(ISBLANK($D{row})),""&$D{row}&"_"&$C{row},"")'.format(row=row)

                row_add_list = [

                    category_item_name,

                    category_question_name,
                    category_name,

                    category_sharedlist,

                    category_label,
                    category_analysisvalue,

                    category_validation,

                    category_validation_field_notblank,
                    category_validation_field_iswholenumber,
                    category_validation_field_isunique,
                    category_validation_field_varnotinexclusions,
                    category_validation_field_isuniquewithinsl,
                    category_validation_field_helper_varnamepluscatname,
                    category_validation_field_helper_varnamepluscatvalue,
                    category_validation_field_helper_slnamepluscatname,
                    category_validation_field_helper_slnamepluscatvalue,
                ]
                row_add = {col[0]: col[1] for col in zip(df.columns,row_add_list)}
                data_add.append(row_add)
                row = row + 1
    
    df = pd.concat([df,pd.DataFrame(data_add)],ignore_index=True)
    df.set_index('Category_Path',inplace=True)
    return df



# def build_mdd_data_validation_df(variables,df,config):
#     return df



def build_variables_df(mdmvariables,mdd_data_variables_df,df_prev,config):
    df = pd.DataFrame(
        data = {
            'Question': [],
            'Question Type': [],
            'Current ShortName': [],
            'Final ShortName': [],
            'Label': [],
            'Include': [],
            'Validation': [],
            'Comment': [],
        },
        index = None,
    )

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
        'report_text_pipein': 'progress populating variables sheet',
    }))
    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue
        question_name = mdmvariable.FullName
        next(performance_counter)
        if not mdmutils.is_data_item(mdmvariable):
            continue
        if mdmutils.is_nocasedata(mdmvariable):
            continue

        question_type = '=VLOOKUP($A{row},\'MDD_Data_Variables\'!$A$2:$G$699999,2,FALSE)'.format(row=row)
        question_shortname_current = '=VLOOKUP($A{row},\'MDD_Data_Variables\'!$A$2:$G$699999,3,FALSE)'.format(row=row)
        question_shortname_final = df_prev.loc[question_name,'Final ShortName'] if question_name in df_prev.index.get_level_values(0) else mdd_data_variables_df.loc[question_name,'ShortName in MDD']
        question_label = '=VLOOKUP($A{row},\'MDD_Data_Variables\'!$A$2:$G$699999,4,FALSE)'.format(row=row)
        question_include_flag = df_prev.loc[question_name,'Include'] if question_name in df_prev.index.get_level_values(0) else ( 'x' if not aa_logic.should_exclude(mdmvariable) else '' )
        question_validation = '=IF(NOT(ISBLANK($F{row})),VLOOKUP($A{row},\'MDD_Data_Variables\'!$A$2:$G$699999,7,FALSE),"")'.format(row=row)
        question_comment = df_prev.loc[question_name,'Comment'] if question_name in df_prev.index.get_level_values(0) else '-'

        row_add_list = [

            question_name,
            
            question_type,
            question_shortname_current,
            question_shortname_final,
            question_label,
            question_include_flag,
            question_validation,
            question_comment,
        ]
        row_add = {col[0]: col[1] for col in zip(df.columns,row_add_list)}
        data_add.append(row_add)
        row = row + 1
    
    df = pd.concat([df,pd.DataFrame(data_add)],ignore_index=True)
    df.set_index('Question',inplace=True)
    return df



def build_analysisvalues_df(mdmvariables,df,mdd_data_categories_prev_df,config):
    df = pd.DataFrame(
        data = {
            'Question': [],
            'ShortName': [],
        },
        index = None,
    )

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
        if mdmutils.is_iterative(mdmvariable) or mdmutils.has_own_categories(mdmvariable):

            categories_data = []
            category_question_shortname = '=VLOOKUP($A{row},MDD_Data_Variables!$A$2:$F$699999,5,FALSE)'.format(row=row)
            category_question_in_exclusions = '=IF(VLOOKUP($A{row},Variables!$A$2:$G$699999,6,FALSE)&""="","(excluded)","")'.format(row=row)

            for col_index_zerobased, mdmcategory in enumerate(mdmutils.list_mdmcategories(mdmvariable)):

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



def build_validationissues_df(variables,df_prev,config):
    df = pd.DataFrame(
        data = {
            'Issues': [],
        },
        index = pd.Series( [], name = 'Where failed' ),
    )
    df.loc['Variables'] = '=_xlfn.XLOOKUP(FALSE,Variables!$G$2:$G$699999,Variables!$A$2:$A$699999,0)'
    df.loc['Categories'] = '=_xlfn.XLOOKUP(FALSE,MDD_Data_Categories!$G$2:$G$699999,MDD_Data_Categories!$A$2:$A$699999,0)'
    return df




def update_map(variables,dataframes,config):
    if not dataframes:
        assert variables, 'update_map: either map or variables are required'
        dataframes = (
            pd.DataFrame( data = { }, index = pd.Series( [], name = 'Description' ), ),
            pd.DataFrame( data = { }, index = pd.Series( [], name = 'Question' ), ),
            pd.DataFrame( data = { }, index = pd.Series( [], name = 'Question' ), ),
            pd.DataFrame( data = { }, index = pd.Series( [], name = 'Where failed' ), ),
            pd.DataFrame( data = { }, index = pd.Series( [], name = 'Question' ), ),
            pd.DataFrame( data = { }, index = pd.Series( [], name = 'Category_Path' ), ),
        )

    overview_df, variables_df, categories_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df = dataframes

    print('building "Overview" sheet...')
    overview_df = build_overview_df(variables,overview_df,config)
    print('building "MDD_Data_Variables" (hidden) sheet...')
    mdd_data_variables_df = build_mdd_data_variables_df(variables,mdd_data_variables_df,config)
    # print('building "MDD_Data_Validation" (hidden) sheet...')
    # mdd_data_validation_df = build_mdd_data_validation_df(variables,mdd_data_validation_df,config)
    print('building "Variables" sheet...')
    variables_df = build_variables_df(variables,mdd_data_variables_df,variables_df,config)
    print('building "Analysis Values" sheet...')
    categories_df = build_analysisvalues_df(variables,categories_df,mdd_data_categories_df,config)
    print('building "MDD_Data_Categories" (hidden) sheet...')
    mdd_data_categories_df = build_mdd_data_categories_df(variables,mdd_data_categories_df,config)
    print('building "Validation Issues Log" sheet...')
    validationissues_df = build_validationissues_df(variables,validationissues_df,config)

    result = (overview_df, variables_df, categories_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df)
    return result



