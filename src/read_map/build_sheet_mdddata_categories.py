
import pandas as pd



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



# from openpyxl.worksheet.formula import ArrayFormula
def build_df(mdmvariables,df_prev,config):
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
        'report_text_pipein': 'progress reading analysis values and populating categories sheet',
    }))
    for mdmvariable in mdmvariables:
        if mdmvariable.Name=='':
            continue
        category_question_name = mdmvariable.FullName
        next(performance_counter)
        if util_mdmvars.is_iterative(mdmvariable) or util_mdmvars.has_own_categories(mdmvariable):
            for mdmsharedlist_name, mdmcategory in util_mdmvars.list_mdmcategories_with_slname(mdmvariable):

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





