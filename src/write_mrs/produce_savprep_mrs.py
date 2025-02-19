import re
import pandas as pd


from . import template
from . import util_mdmvars










class ParseMap:
    # actually, it's tricky
    # reading excel is harder than it looks

    # let's say, we have 3 columns
    # - Shortname (MDD)
    # - SHortname (Prev)
    # - Shortname (Final)

    # the column "Final" contains a formula
    # if Excel was edited, we should read from thet "Final" column
    # but the Excel was just saved by a tool and not opened by a user,
    # the value in "Final" column is not calculated
    # we can still read from "prev" column

    CONFIG_MAP_VARS_INCLUDE_COLUMN_FORMULA = 'Include (Final)'
    CONFIG_MAP_VARS_INCLUDE_COLUMN_PREV = 'Include (read from prev Map)'
    CONFIG_MAP_USERVARS_INCLUDE_COLUMN = 'Include'
    CONFIG_MAP_VARS_SHORTNAME_COLUMN_FORMULA = 'ShortName (Final)'
    CONFIG_MAP_VARS_SHORTNAME_COLUMN_PREV = 'ShortName (prev Map)'
    CONFIG_MAP_USERVARS_SHORTNAME_COLUMN = 'ShortName'
    CONFIG_MAP_VARS_LABEL_COLUMN_FORMULA = 'Label (Final)'
    CONFIG_MAP_VARS_LABEL_COLUMN_PREV = 'Label (prev Map)'
    CONFIG_MAP_USERVARS_LABEL_COLUMN = 'Label'
    CONFIG_MAP_VARS_ISDATAVAR_COLUMN = 'Question - Is data variable'
    CONFIG_MAP_CATS_ANALYSISVALUE_COLUMN_FORMULA = 'Analysis Value (Final)'
    CONFIG_MAP_CATS_ANALYSISVALUE_COLUMN_PREV = 'Analysis Value (prev Map)'
    CONFIG_MAP_CATS_LABEL_COLUMN_FORMULA = 'Label (Final)'
    CONFIG_MAP_CATS_LABEL_COLUMN_PREV = 'Label (prev Map)'

    def __init__(self,dataframes):
        assert not not dataframes, 'Error: Map is required but it missing'

        overview_df, variables_df, analysisvalues_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df = dataframes
        self.overview_df = overview_df
        self.variables_df = variables_df
        self.analysisvalues_df = analysisvalues_df
        self.validationissues_df = validationissues_df
        self.mdd_data_variables_df = mdd_data_variables_df
        self.mdd_data_categories_df = mdd_data_categories_df

        assert mdd_data_variables_df is not None, 'Error: \'MDD Variables\' sheet is required but it missing'
        assert mdd_data_categories_df is not None, 'Error: \'MDD Categories\' sheet is required but it missing'
    
    def read_question_is_included(self,question_name):
        assert question_name in self.mdd_data_variables_df.index.get_level_values(0), 'Error: please reload the map; "{var}" variable is found in MDD but it missing in the map. Please make sure the Map is up to date, reload the Map first'.format(var=question_name)
        if question_name in self.variables_df.index.get_level_values(0):
            value = self.variables_df.loc[question_name,self.CONFIG_MAP_USERVARS_INCLUDE_COLUMN]
            value = not not value # convert "x" to True - uniform looking data type, so that does not surprise me later
        else:
            value = self.mdd_data_variables_df.loc[question_name,self.CONFIG_MAP_VARS_INCLUDE_COLUMN_FORMULA]
            if not self.has_value_numeric(value):
                value = self.mdd_data_variables_df.loc[question_name,self.CONFIG_MAP_VARS_INCLUDE_COLUMN_PREV]
        return value

    def read_question_shortname(self,question_name):
        assert question_name in self.mdd_data_variables_df.index.get_level_values(0), 'Error: please reload the map; "{var}" variable is found in MDD but it missing in the map. Please make sure the Map is up to date, reload the Map first'.format(var=question_name)
        if question_name in self.variables_df.index.get_level_values(0):
            value = self.variables_df.loc[question_name,self.CONFIG_MAP_USERVARS_SHORTNAME_COLUMN]
        else:
            value = self.mdd_data_variables_df.loc[question_name,self.CONFIG_MAP_VARS_SHORTNAME_COLUMN_FORMULA]
            if not self.has_value_text(value):
                value = self.mdd_data_variables_df.loc[question_name,self.CONFIG_MAP_VARS_SHORTNAME_COLUMN_PREV]
        return value

    def read_question_label(self,question_name):
        if question_name in self.variables_df.index.get_level_values(0):
            value = self.variables_df.loc[question_name,self.CONFIG_MAP_USERVARS_LABEL_COLUMN]
        else:
            value = self.mdd_data_variables_df.loc[question_name,self.CONFIG_MAP_VARS_LABEL_COLUMN_FORMULA]
            if not self.has_value_text(value):
                value = self.mdd_data_variables_df.loc[question_name,self.CONFIG_MAP_VARS_LABEL_COLUMN_PREV]
        return value

    def read_category_analysisvalue(self,question_name,category_name):
        value = None
        found_user_input = False
        category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
        if question_name in self.analysisvalues_df.index.get_level_values(0):
            row_number = self.analysisvalues_df.index.get_loc(question_name)
            row_category_names = self.analysisvalues_df.iloc[row_number,1:].tolist()
            if category_name in row_category_names:
                # row_category_labels = self.analysisvalues_df.iloc[row_number+1,1:].tolist()
                row_category_analysisvalues = self.analysisvalues_df.iloc[row_number+2,1:].tolist()
                # row_category_validation = self.analysisvalues_df.iloc[row_number+3,1:].tolist()
                index = row_category_names.index(category_name)
                value = row_category_analysisvalues[index]
                found_user_input = True
        if not found_user_input:
            value = self.mdd_data_categories_df.loc[category_path,self.CONFIG_MAP_CATS_ANALYSISVALUE_COLUMN_FORMULA]
            if not self.has_value_numeric(value):
                value = self.mdd_data_categories_df.loc[category_path,self.CONFIG_MAP_CATS_ANALYSISVALUE_COLUMN_PREV]
        if not self.has_value_numeric(value):
            return None
        else:
            return value

    def read_category_label(self,question_name,category_name):
        value = None
        found_user_input = False
        category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
        if question_name in self.analysisvalues_df.index.get_level_values(0):
            row_number = self.analysisvalues_df.index.get_loc(question_name)
            row_category_names = self.analysisvalues_df.iloc[row_number,1:].tolist()
            if category_name in row_category_names:
                row_category_labels = self.analysisvalues_df.iloc[row_number+1,1:].tolist()
                # row_category_analysisvalues = self.analysisvalues_df.iloc[row_number+2,1:].tolist()
                # row_category_validation = self.analysisvalues_df.iloc[row_number+3,1:].tolist()
                index = row_category_names.index(category_name)
                value = row_category_labels[index]
                found_user_input = True
        if not found_user_input:
            value = self.mdd_data_categories_df.loc[category_path,self.CONFIG_MAP_CATS_LABEL_COLUMN_FORMULA]
            if not self.has_value_numeric(value):
                value = self.mdd_data_categories_df.loc[category_path,self.CONFIG_MAP_CATS_LABEL_COLUMN_PREV]
        if not self.has_value_numeric(value):
            return None
        else:
            return value

    def has_value_numeric(self,arg):
        if pd.isna(arg):
            return False
        if arg is None:
            return False
        if arg==0:
            return True
        if arg == False:
            return True # false evaluates to 0 which is numeric
        if arg=='':
            return False
        return not not arg

    def has_value_text(self,arg):
        if pd.isna(arg):
            return False
        if arg is None:
            return False
        if arg==0:
            return True
        if arg == False:
            return False
        if arg=='':
            return False
        return not not arg



def sanitize_txt_escape_quotes(txt):
    t = '{t}'.format(t=txt)
    t = t.replace('"','""')
    t = re.sub(r'\s*\n\s*',' ',t,flags=re.I|re.DOTALL)
    return t



def sanitize_analysisvalue(val):
    if val is None:
        return ''
    val = float(val)
    val_rounded = int(round(val))
    if abs(val-val_rounded)<0.001:
        val = val_rounded
    return val_rounded



def make_path_to_field(variable):
    if util_mdmvars.is_helper_field(variable):
        return make_path_to_field(variable.Parent) + '.HelperFields["{name}"]'.format(name=variable.Name)
    if '@class' in variable.Name:
        return make_path_to_field(variable.Parent)
    # path = re.replace(r'\[[^\]]*?\]','',variable.FullName,flags=re.I|re.DOTALL)
    # parts = path.split('.')
    # return 'mdm' + ''.join(['.Fields["{name}"]'.format(name=part) for part in parts])
    if util_mdmvars.has_parent(variable):
        return make_path_to_field(variable.Parent) + '.Fields["{name}"]'.format(name=variable.Name)
    else:
        return 'mDocument' + '.Fields["{name}"]'.format(name=variable.Name)



def generate_savprep_mrs_template(config):
    result = template.TEMPLATE
    result = result.replace('{{{addin_file_path}}}',config['out_filename_include_addin_filenamepart'])
    return result




def generate_savprep_mrs_include_addin(mddvariables,dataframes,config):
    assert not not mddvariables, 'Error: MDD is required but it missing'
    map = ParseMap(dataframes)

    result = ''

    for variable in mddvariables:
        if variable.Name=='':
            continue
        question_name = variable.FullName

        is_included            = map.read_question_is_included(question_name)
        is_data_variable       = util_mdmvars.is_data_item(variable) # map.read_question_...()
        is_iterative           = util_mdmvars.is_iterative(variable) or util_mdmvars.has_own_categories(variable)
        question_shortname     = map.read_question_shortname(question_name)
        question_label         = map.read_question_label(question_name)

        result = result + '\n'
        result = result + '\' {name}\n'.format(name=question_name)
        result = result + 'debug.echo("processing {name}...")\n'.format(name=question_name)
        result = result + 'debug.log("processing {name}...")\n'.format(name=question_name)
        result = result + 'set oField = {address_field}\n'.format(address_field=make_path_to_field(variable))

        if is_iterative:
            for category in util_mdmvars.list_mdmcategories(variable):
                category_name = category.Name

                category_analysisvalue = map.read_category_analysisvalue(question_name,category_name)
                category_label = map.read_category_label(question_name,category_name)

                result = result + '{apostrophe}set oCategory = oField.Categories["{catname}"]\n'.format(catname=category_name,apostrophe='' if category_analysisvalue is not None else '\'')
                result = result + '{apostrophe}oCategory.Label = "{val}"\n'.format(val=sanitize_txt_escape_quotes(category_label),apostrophe='' if category_analysisvalue is not None else '\'')
                result = result + '{apostrophe}oCategory.Properties.Item["Value"] = {val}\n'.format(val=sanitize_analysisvalue(category_analysisvalue),apostrophe='' if category_analysisvalue is not None else '\'')

        if is_data_variable:
            if is_included:
                result = result + 'oField.Label = "{txt}"\n'.format(txt=sanitize_txt_escape_quotes(question_label))
                result = result + 'SetAlias(oField,"{prop}")\n'.format(prop=sanitize_txt_escape_quotes(question_shortname))
            else:
                result = result + 'oField.HasCaseData = False\n'



    return result
