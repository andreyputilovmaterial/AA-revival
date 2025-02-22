
import pandas as pd




from . import build_sheet_overview
from . import build_sheet_variables
from . import build_sheet_analysisvalues
from . import build_sheet_mdddata_variables
from . import build_sheet_mdddata_categories
from . import build_sheet_validation

from . import excel_format_sheet

from .column_definitions import sheet_overview
from .column_definitions import sheet_mdddata_variables
from .column_definitions import sheet_mdddata_categories
from .column_definitions import sheet_variables
from .column_definitions import sheet_analysisvalues
from .column_definitions import sheet_validation




def has_value_numeric(arg):
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

def has_value_text(arg):
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






class Map:
    def __init__(self,map_filename=None,config={}):

        if map_filename:
            with pd.ExcelFile(map_filename,engine='openpyxl') as xls:
                overview_df = xls.parse(sheet_name=sheet_overview.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                variables_df = xls.parse(sheet_name=sheet_variables.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                analysisvalues_df = xls.parse(sheet_name=sheet_analysisvalues.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                # analysisvalues_df = xls.parse(sheet_name=sheet_analysisvalues.sheet_name, header=0, index_col=0, dtype=str, keep_default_na=False).fillna('')
                # analysisvalues_df = analysisvalues_df.convert_dtypes(convert_boolean=False)
                validationissues_df = xls.parse(sheet_name=sheet_validation.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                mdd_data_variables_df = xls.parse(sheet_name=sheet_mdddata_variables.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                mdd_data_categories_df = xls.parse(sheet_name=sheet_mdddata_categories.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
        else:
            overview_df, variables_df, analysisvalues_df, validationissues_df, mdd_data_variables_df, mdd_data_categories_df = ( None, None, None, None, None, None, )

        self.overview_df = overview_df
        self.variables_df = variables_df
        self.analysisvalues_df = analysisvalues_df
        self.validationissues_df = validationissues_df
        self.mdd_data_variables_df = mdd_data_variables_df
        self.mdd_data_categories_df = mdd_data_categories_df

        self.config = config



    def get_mdd_path(self):
        df =self.overview_df
        try:
            mdd_filename = df.at['MDD path','Value']
        except Exception as e:
            raise FileNotFoundError('Can\'t read MDD file name from map: {e}'.format(e=e)) from e
        return mdd_filename
    


    def update(self,mdd):
        assert mdd, 'update_map: mdd is required but it missing'

        config = self.config

        overview_df = self.overview_df
        variables_df = self.variables_df
        analysisvalues_df = self.analysisvalues_df
        validationissues_df = self.validationissues_df
        mdd_data_variables_df = self.mdd_data_variables_df
        mdd_data_categories_df = self.mdd_data_categories_df

        print('building "Overview" sheet...')
        overview_df                 = build_sheet_overview.build_df(mdd,config)
        print('building "MDD_Data_Variables" (hidden) sheet...')
        mdd_data_variables_df       = build_sheet_mdddata_variables.build_df(mdd,self,config)
        print('building "Variables" sheet...')
        variables_df                = build_sheet_variables.build_df(mdd,self,config)
        print('building "MDD_Data_Categories" (hidden) sheet...')
        mdd_data_categories_df      = build_sheet_mdddata_categories.build_df(mdd,self,config)
        print('building "Analysis Values" sheet...')
        analysisvalues_df               = build_sheet_analysisvalues.build_df(mdd,self,config)
        print('building "Validation Issues Log" sheet...')
        validationissues_df         = build_sheet_validation.build_df(config)
        
        result = Map(None,config=self.config)

        result.overview_df = overview_df
        result.variables_df = variables_df
        result.analysisvalues_df = analysisvalues_df
        result.validationissues_df = validationissues_df
        result.mdd_data_variables_df = mdd_data_variables_df
        result.mdd_data_categories_df = mdd_data_categories_df

        return result



    def write_to_file(self,out_filename):

        config = self.config

        overview_df = self.overview_df
        variables_df = self.variables_df
        analysisvalues_df = self.analysisvalues_df
        validationissues_df = self.validationissues_df
        mdd_data_variables_df = self.mdd_data_variables_df
        mdd_data_categories_df = self.mdd_data_categories_df

        
        # excel_format_sheet.format_sheet_overview(overview_df)
        # excel_format_sheet.format_sheet_variables(variables_df)
        # excel_format_sheet.format_sheet_analysisvalues(analysisvalues_df)
        with pd.ExcelWriter(out_filename, engine='openpyxl') as writer:
            overview_df.to_excel(writer, sheet_name=sheet_overview.sheet_name)
            variables_df.to_excel(writer, sheet_name=sheet_variables.sheet_name)
            analysisvalues_df.to_excel(writer, sheet_name=sheet_analysisvalues.sheet_name)
            validationissues_df.to_excel(writer, sheet_name=sheet_validation.sheet_name)
            mdd_data_variables_df.to_excel(writer, sheet_name=sheet_mdddata_variables.sheet_name)
            mdd_data_categories_df.to_excel(writer, sheet_name=sheet_mdddata_categories.sheet_name)
            # mdd_data_validation_df.to_excel(writer, sheet_name='ValidationVariables')
            excel_format_sheet.format_sheet_overview(writer.sheets[sheet_overview.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_variables.sheet_name])
            excel_format_sheet.format_sheet_variables(writer.sheets[sheet_variables.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_analysisvalues.sheet_name])
            excel_format_sheet.format_sheet_analysisvalues(writer.sheets[sheet_analysisvalues.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_validation.sheet_name])
            excel_format_sheet.format_sheet_validation(writer.sheets[sheet_validation.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_mdddata_variables.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_mdddata_categories.sheet_name])
            # excel_format_sheet.format_with_stdstyles(writer.sheets['ValidationVariables'])
            writer.sheets[sheet_mdddata_variables.sheet_name].protection.sheet = True # openpyxl-specific syntax
            writer.sheets[sheet_mdddata_variables.sheet_name].sheet_state = 'hidden' # openpyxl-specific syntax
            writer.sheets[sheet_mdddata_categories.sheet_name].protection.sheet = True # openpyxl-specific syntax
            writer.sheets[sheet_mdddata_categories.sheet_name].sheet_state = 'hidden' # openpyxl-specific syntax
            # writer.sheets['ValidationVariables'].protection.sheet = True # openpyxl-specific syntax
            # writer.sheets['ValidationVariables'].sheet_state = 'hidden' # openpyxl-specific syntax
            # writer.sheets[sheet_mdddata_variables.sheet_name].protect()
            # writer.sheets[sheet_mdddata_categories.sheet_name].protect()
            # writer.sheets['ValidationVariables'].protect()
            # writer.sheets[sheet_mdddata_variables.sheet_name].hide()
            # writer.sheets[sheet_mdddata_categories.sheet_name].hide()
            # writer.sheets['ValidationVariables'].hide()




    def read_question_is_included(self,question_name):
        assert question_name in self.mdd_data_variables_df.index.get_level_values(0), 'Error: please reload the map; "{var}" variable is found in MDD but it missing in the map. Please make sure the Map is up to date, reload the Map first'.format(var=question_name)
        if question_name in self.variables_df.index.get_level_values(0):
            value = self.variables_df.loc[question_name,sheet_variables.column_names['col_include']]
            value = not not value # convert "x" to True - uniform looking data type, so that does not surprise me later
        else:
            value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_include']]
            if not has_value_numeric(value):
                value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_include_prev']]
        return value

    def read_question_shortname(self,question_name):
        assert question_name in self.mdd_data_variables_df.index.get_level_values(0), 'Error: please reload the map; "{var}" variable is found in MDD but it missing in the map. Please make sure the Map is up to date, reload the Map first'.format(var=question_name)
        if question_name in self.variables_df.index.get_level_values(0):
            value = self.variables_df.loc[question_name,sheet_variables.column_names['col_shortname']]
        else:
            value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_shortname']]
            if not has_value_text(value):
                value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_shortname_prev']]
        return value

    def read_question_label(self,question_name):
        if question_name in self.variables_df.index.get_level_values(0):
            value = self.variables_df.loc[question_name,sheet_variables.column_names['col_label']]
        else:
            value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_label']]
            if not has_value_text(value):
                value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_label_prev']]
        return value

    def read_question_comment(self,question_name):
        if question_name in self.variables_df.index.get_level_values(0):
            value = self.variables_df.loc[question_name,sheet_variables.column_names['col_comment']]
        else:
            value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_comment']]
            if not has_value_text(value):
                value = self.mdd_data_variables_df.loc[question_name,sheet_mdddata_variables.column_names['col_comment_prev']]
        # if value==0 or value=='0': # dirty, sorry: suppressing some artifacts from vlookup
        #     return ''
        return value

    def read_category_analysisvalue(self,question_name,category_name):
        value = None
        found_user_input = False
        category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
        if question_name in self.analysisvalues_df.index.get_level_values(0):
            row_number = self.analysisvalues_df.index.get_loc(question_name)
            row_category_names = self.analysisvalues_df.iloc[row_number,1:].tolist()
            if category_name in row_category_names:
                # an issue is happening here
                # pandas keeps reading 1's as True, and I can't find a workaround
                # I am trying to wrap the row with a call to Series(dtype=int), I am trying to call .astype(int)
                # I am trying to set dtypes=str when reading Excel
                # I am trying to call df = df.convert_dtypes(convert_boolean=False) right after opening the file
                # no luck
                # I just edited Excel and I know I entered 1
                # I add a breakpoint here, and I see True, not 1, whatever I do
                # crazy

                # row_category_labels = self.analysisvalues_df.iloc[row_number+1,1:].tolist()
                row_category_analysisvalues = pd.Series(self.analysisvalues_df.iloc[row_number+2,1:],dtype=int).tolist()
                # row_category_analysisvalues = pd.Series(self.analysisvalues_df.iloc[row_number+2,1:],dtype=int).astype(int).tolist()
                # row_category_validation = self.analysisvalues_df.iloc[row_number+3,1:].tolist()
                index = row_category_names.index(category_name)
                value = row_category_analysisvalues[index]
                if value==True:
                    value = 1
                elif value==False:
                    value = 0
                found_user_input = True
        if not found_user_input:
            value = self.mdd_data_categories_df.loc[category_path,sheet_mdddata_categories.column_names['col_analysisvalue']]
            if not has_value_numeric(value):
                value = self.mdd_data_categories_df.loc[category_path,sheet_mdddata_categories.column_names['col_analysisvalue_prev']]
        if not has_value_numeric(value):
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
            value = self.mdd_data_categories_df.loc[category_path,sheet_mdddata_categories.column_names['col_label']]
            if not has_value_text(value):
                value = self.mdd_data_categories_df.loc[category_path,sheet_mdddata_categories.column_names['col_label_prev']]
        if not has_value_text(value):
            return None
        else:
            return value
