


sheet_name = 'Variables'

columns = [
    'Question',
    'Question Type',
    'Label',
    'ShortName',
    'Include',
    'Validation',
    'Comment',
]

column_aliases = [
    'col_question',
    'col_question_type',
    'col_label',
    'col_shortname',
    'col_include',
    'col_validation',
    'col_comment',
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
