


sheet_name = 'Validation Issues Log'

columns = [
    'Where failed',
    'Issues',
]

column_aliases = [
    'col_path',
    'col_issues',
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
