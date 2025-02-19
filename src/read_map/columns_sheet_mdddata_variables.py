


sheet_name = 'MDD_Data_Variables'

columns = [
    'Question',
    'Question Type',
    'Question - Is data variable',
    'Label (in MDD)',
    'Label (prev Map)',
    'Label (Final)',
    'ShortName (in MDD)',
    'ShortName (prev Map)',
    'ShortName (Final)',
    'Comment (prev Map)',
    'Comment (Final)',
    'Include (auto-detected from MDD)',
    'Include (read from prev Map)',
    'Include (Final)',
    'Validation',
    'Helper Validation Field - ShortName Lowercase',
    'Helper Validation - Not blank',
    'Helper Validation - Is alphanumeric',
    'Helper Validation - Does not start with a number',
    'Helper Validation - Is unique',
]

column_aliases = [
    'col_question',
    'col_question_type',
    'col_question_isdatavariable',
    'col_label_mdd',
    'col_label_prev',
    'col_label',
    'col_shortname_mdd',
    'col_shortname_prev',
    'col_shortname',
    'col_comment_prev',
    'col_comment',
    'col_include_mdd',
    'col_include_prev',
    'col_include',
    'col_validation',
    'col_helper_shortname_clean_lcase',
    'col_helper_validation_notblank',
    'col_helper_validation_isalphanumeric',
    'col_helper_validation_notstartswithnumber',
    'col_helper_validation_isunique',
]

column_names = {
    prop: value for prop, value in zip( column_aliases, columns )
}

column_letters = {
    prop: value for prop, value in zip( column_aliases, ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV','AW','AX','AY','AZ',][:len(columns)] )
}

column_vlookup_index = {
    col_alias: index+1 for index ,col_alias in enumerate( column_aliases )
}
