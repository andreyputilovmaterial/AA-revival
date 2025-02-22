


sheet_name = 'MDD_Data_Categories'

columns = [
    'Path',
    'Question',
    'Category',
    'Shared List',
    'Label (in MDD)',
    'Label (prev Map)',
    'Label (Final)',
    'Analysis Value (in MDD)',
    'Analysis Value (prev Map)',
    'Analysis Value (Final)',
    'Validation',
    'Helper Validation - Var not in exclusions',
    'Helper Validation - Not blank',
    'Helper Validation - Is whole number',
    'Helper Validation - Is unique',
    'Helper Validation - Is unique within shared list',
    'Helper Validation Field - Question + Category Name',
    'Helper Validation Field - Question + Category Value',
    'Helper Validation Field - SL Name + Category Name',
    'Helper Validation Field - SL Name + Category Value',
]

column_aliases = [
    'col_path',
    'col_question',
    'col_category',
    'col_sharedlist',
    'col_label_mdd',
    'col_label_prev',
    'col_label',
    'col_analysisvalue_mdd',
    'col_analysisvalue_prev',
    'col_analysisvalue',
    'col_validation',
    'col_helper_validation_varnotinexclusions',
    'col_helper_validation_notblank',
    'col_helper_validation_iswhole',
    'col_helper_validation_isunique',
    'col_helper_validation_isuniquewithinsl',
    'col_helper_field_categorypath',
    'col_helper_field_categoryvalue',
    'col_helper_field_slpath',
    'col_helper_field_slvalue',
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
