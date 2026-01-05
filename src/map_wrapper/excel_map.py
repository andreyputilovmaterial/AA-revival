
import pandas as pd
import numpy as np # this is almost unnecessary but I am checking types for debugging and validation

import sys # for error reporting, to print to stderr



from .backup_manager import BackupManager

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








class ExcelMap:
    
    class ExcelMapException(Exception):
        """General superclass for ExcelMap exceptions"""

    class CellNotFound(ExcelMapException):
        """Cell not found"""
    
    
    
    # when  mao object is created, it is not filled with values
    # it can be created providing some filename that it should load, or not, than a blank new file will be the map
    # "update" method should be called to have the map filled
    def __init__(self,map_filename=None,config={}):

        if map_filename:
            with pd.ExcelFile(map_filename,engine='openpyxl') as xls:
                df_overview = xls.parse(sheet_name=sheet_overview.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                df_userinput_variables = xls.parse(sheet_name=sheet_variables.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                df_userinput_analysisvalues = xls.parse(sheet_name=sheet_analysisvalues.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                # df_userinput_analysisvalues = xls.parse(sheet_name=sheet_analysisvalues.sheet_name, header=0, index_col=0, dtype=str, keep_default_na=False).fillna('')
                # df_userinput_analysisvalues = df_userinput_analysisvalues.convert_dtypes(convert_boolean=False)
                df_validation = xls.parse(sheet_name=sheet_validation.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                df_mdddata_variables = xls.parse(sheet_name=sheet_mdddata_variables.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
                df_mdddata_categories = xls.parse(sheet_name=sheet_mdddata_categories.sheet_name, header=0, index_col=0, keep_default_na=False).fillna('')
        else:
            df_overview, df_userinput_variables, df_userinput_analysisvalues, df_validation, df_mdddata_variables, df_mdddata_categories = ( None, None, None, None, None, None, )

        self.df_overview = df_overview
        self.df_userinput_variables = df_userinput_variables
        self.df_userinput_analysisvalues = df_userinput_analysisvalues
        self.df_validation = df_validation
        self.df_mdddata_variables = df_mdddata_variables
        self.df_mdddata_categories = df_mdddata_categories

        self.config = config



    # helper method to have MDD path fetched from the map, stored on "Overview" sheet
    def get_mdd_path(self):
        df = self.df_overview
        try:
            mdd_filename = df.at['MDD path','Value']
        except Exception as e:
            raise FileNotFoundError('Can\'t read MDD file name from map: {e}'.format(e=e)) from e
        return mdd_filename
    


    # when  mao object is created, it is not filled with values
    # it can be created providing some filename that it should load, or not, than a blank new file will be the map
    # "update" method should be called to have the map filled
    def update(self,mdd):
        assert mdd, 'update_map: mdd is required but it missing'

        config = self.config

        result = self.__class__(None,config=self.config)
        result.df_overview = self.df_overview
        result.df_mdddata_variables = self.df_mdddata_variables
        result.df_mdddata_categories = self.df_mdddata_categories
        result.df_userinput_variables = self.df_userinput_variables
        result.df_userinput_analysisvalues = self.df_userinput_analysisvalues
        result.df_validation = self.df_validation

        print('building "Overview" sheet...')
        result.df_overview                 = build_sheet_overview.build_df(mdd,config)
        print('building "MDD_Data_Variables" (hidden) sheet...')
        result.df_mdddata_variables        = build_sheet_mdddata_variables.build_df(mdd,result,config)
        print('building "MDD_Data_Categories" (hidden) sheet...')
        result.df_mdddata_categories       = build_sheet_mdddata_categories.build_df(mdd,result,config)
        print('building "Variables" sheet...')
        result.df_userinput_variables      = build_sheet_variables.build_df(mdd,result,config)
        print('building "Analysis Values" sheet...')
        result.df_userinput_analysisvalues = build_sheet_analysisvalues.build_df(mdd,result,config)
        print('building "Validation Issues Log" sheet...')
        result.df_validation               = build_sheet_validation.build_df(config)
        
        return result
    


    def prepopulate_analysis_values(self):
        print('pre-populating the "Analysis Values" sheet...')
        self.df_userinput_analysisvalues = build_sheet_analysisvalues.prepopulate(self.df_userinput_analysisvalues,self.df_mdddata_categories,self.config)
        return self.df_userinput_analysisvalues



    def write_to_file(self,out_filename):

        config = self.config

        df_overview = self.df_overview
        df_userinput_variables = self.df_userinput_variables
        df_userinput_analysisvalues = self.df_userinput_analysisvalues
        df_validation = self.df_validation
        df_mdddata_variables = self.df_mdddata_variables
        df_mdddata_categories = self.df_mdddata_categories

        
        with pd.ExcelWriter(out_filename, engine='openpyxl') as writer:
            df_overview.to_excel(writer, sheet_name=sheet_overview.sheet_name)
            df_userinput_variables.to_excel(writer, sheet_name=sheet_variables.sheet_name)
            df_userinput_analysisvalues.to_excel(writer, sheet_name=sheet_analysisvalues.sheet_name)
            df_validation.to_excel(writer, sheet_name=sheet_validation.sheet_name)
            df_mdddata_variables.to_excel(writer, sheet_name=sheet_mdddata_variables.sheet_name)
            df_mdddata_categories.to_excel(writer, sheet_name=sheet_mdddata_categories.sheet_name)
            excel_format_sheet.format_sheet_overview(writer.sheets[sheet_overview.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_variables.sheet_name])
            excel_format_sheet.format_sheet_variables(writer.sheets[sheet_variables.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_analysisvalues.sheet_name])
            excel_format_sheet.format_sheet_analysisvalues(writer.sheets[sheet_analysisvalues.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_validation.sheet_name])
            excel_format_sheet.format_sheet_validation(writer.sheets[sheet_validation.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_mdddata_variables.sheet_name])
            excel_format_sheet.format_with_stdstyles(writer.sheets[sheet_mdddata_categories.sheet_name])
            writer.sheets[sheet_mdddata_variables.sheet_name].protection.sheet = True # openpyxl-specific syntax
            writer.sheets[sheet_mdddata_variables.sheet_name].sheet_state = 'hidden' # openpyxl-specific syntax
            writer.sheets[sheet_mdddata_categories.sheet_name].protection.sheet = True # openpyxl-specific syntax
            writer.sheets[sheet_mdddata_categories.sheet_name].sheet_state = 'hidden' # openpyxl-specific syntax


    @staticmethod
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

    @staticmethod
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

    def read_question_is_included_column_userinput(self,question_name):
        if self.df_userinput_variables is None or question_name not in self.df_userinput_variables.index.get_level_values(0):
            raise self.CellNotFound('"is_included" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_variables.sheet_name))
        value = self.df_userinput_variables.loc[question_name,sheet_variables.column_names['col_include']]
        value = not not value # convert "x" to True - uniform looking data type, so that does not surprise me later
        return value

    def read_question_is_included_column_historic(self,question_name):
        if self.df_mdddata_variables is None or question_name not in self.df_mdddata_variables.index.get_level_values(0):
            raise self.CellNotFound('"is_included_prev" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        # value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_include']]
        # if not self.has_value_numeric(value):
        value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_include_prev']]
        return value
        
    def read_question_is_included_column_mdd(self,question_name):
        if self.df_mdddata_variables is None or question_name not in self.df_mdddata_variables.index.get_level_values(0):
            raise self.CellNotFound('"is_included_mdd" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_include_mdd']]
        return value
        
    def read_question_is_included(self,question_name):
        try:
            return self.read_question_is_included_column_userinput(question_name)
        except self.CellNotFound:
            value = self.read_question_is_included_column_historic(question_name)
            # if not self.has_value_numeric(value):
            #     value = self.read_question_is_included_column_mdd(question_name)
            return value

    def read_question_shortname_column_userinput(self,question_name):
        if self.df_userinput_variables is None or question_name not in self.df_userinput_variables.index.get_level_values(0):
            raise self.CellNotFound('"shortname" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_variables.sheet_name))
        value = self.df_userinput_variables.loc[question_name,sheet_variables.column_names['col_shortname']]
        return value

    def read_question_shortname_column_historic(self,question_name):
        if self.df_mdddata_variables is None or question_name not in self.df_mdddata_variables.index.get_level_values(0):
            raise self.CellNotFound('"shortname_prev" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        # value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_shortname']]
        # if not self.has_value_text(value):
        value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_shortname_prev']]
        return value
        
    def read_question_shortname_column_mdd(self,question_name):
        if self.df_mdddata_variables is None or question_name not in self.df_mdddata_variables.index.get_level_values(0):
            raise self.CellNotFound('"shortname_mdd" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_shortname_mdd']]
        return value
        
    def read_question_shortname(self,question_name):
        try:
            return self.read_question_shortname_column_userinput(question_name)
        except self.CellNotFound:
            value = self.read_question_shortname_column_historic(question_name)
            # if not self.has_value_text(value):
            #     value = self.read_question_shortname_column_mdd(question_name)
            return value

    def read_question_label_column_userinput(self,question_name):
        if self.df_userinput_variables is None or question_name not in self.df_userinput_variables.index.get_level_values(0):
            raise self.CellNotFound('"label" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_variables.sheet_name))
        value = self.df_userinput_variables.loc[question_name,sheet_variables.column_names['col_label']]
        return value

    def read_question_label_column_historic(self,question_name):
        if self.df_mdddata_variables is None or question_name not in self.df_mdddata_variables.index.get_level_values(0):
            raise self.CellNotFound('"label_prev" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        # value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_label']]
        # if not self.has_value_text(value):
        value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_label_prev']]
        return value
        
    def read_question_label_column_mdd(self,question_name):
        if self.df_mdddata_variables is None or question_name not in self.df_mdddata_variables.index.get_level_values(0):
            raise self.CellNotFound('"label_mdd" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_label_mdd']]
        return value
        
    def read_question_label(self,question_name):
        try:
            return self.read_question_label_column_userinput(question_name)
        except self.CellNotFound:
            value = self.read_question_label_column_historic(question_name)
            # if not self.has_value_text(value):
            #     value = self.read_question_label_column_mdd(question_name)
            return value

    def read_question_comment_column_userinput(self,question_name):
        if self.df_userinput_variables is None or question_name not in self.df_userinput_variables.index.get_level_values(0):
            raise self.CellNotFound('"comment" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_variables.sheet_name))
        value = self.df_userinput_variables.loc[question_name,sheet_variables.column_names['col_comment']]
        return value

    def read_question_comment_column_historic(self,question_name):
        if self.df_mdddata_variables is None or question_name not in self.df_mdddata_variables.index.get_level_values(0):
            raise self.CellNotFound('"comment_prev" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        # value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_comment']]
        # if not self.has_value_text(value):
        value = self.df_mdddata_variables.loc[question_name,sheet_mdddata_variables.column_names['col_comment_prev']]
        return value
        
    def read_question_comment(self,question_name):
        try:
            return self.read_question_comment_column_userinput(question_name)
        except self.CellNotFound:
            value = self.read_question_comment_column_historic(question_name)
            return value


    def read_category_analysisvalue_column_userinput(self,question_name,category_name):
        def get_loc(question_name):
            # additional code to have better understanding of what happens when reading analysisvalues sheet
            # sorry, looks ugly a bit, 
            row_number = None
            _possible_pandas_error = None
            try:
                row_number = self.df_userinput_analysisvalues.index.get_loc(question_name)
            except Exception as e:
                _possible_pandas_error = e
            if _possible_pandas_error:
                if isinstance(_possible_pandas_error,KeyError):
                    raise self.__class__.ExcelMapException('ERROR: Reading "Analysis Values" sheet: can\'t find row index for "{f}" variable'.format(f=question_name))
                else:
                    raise _possible_pandas_error
            elif np.ndim(row_number)>0:
                raise self.__class__.ExcelMapException('ERROR: Reading "Analysis Values" sheet: Something is off, multiple entries for "{f}" variable'.format(f=question_name))
            else:
                pass # all good, continue
            return row_number
        try:
            value = None
            category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
            if self.df_userinput_analysisvalues is not None and question_name in self.df_userinput_analysisvalues.index.get_level_values(0):
                row_number = get_loc(question_name)
                assert np.isscalar(row_number), 'ERROR: something is off, from the excel with map: trying to find index of row "{row}" on "Analysis Values" sheet: row_number must be scalar numeric, got {var_type}'.format(var_type=type(row_number),row=question_name)
                row_category_names = self.df_userinput_analysisvalues.iloc[row_number,1:].squeeze().tolist()
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

                    # row_category_labels = self.df_userinput_analysisvalues.iloc[row_number+1,1:].squeeze().tolist()
                    row_category_analysisvalues = pd.Series(self.df_userinput_analysisvalues.iloc[row_number+2,1:],dtype=int).squeeze().tolist()
                    # row_category_analysisvalues = pd.Series(self.df_userinput_analysisvalues.iloc[row_number+2,1:],dtype=int).astype(int).squeeze().tolist()
                    # row_category_validation = self.df_userinput_analysisvalues.iloc[row_number+3,1:].squeeze().tolist()
                    index = row_category_names.index(category_name)
                    value = row_category_analysisvalues[index]
                    if value==True:
                        value = 1
                    elif value==False:
                        value = 0
                    return value
            raise self.CellNotFound('"CAT_XXX" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=category_path,sheet_name=sheet_analysisvalues.sheet_name))
        except Exception as e:
            print('ERROR: something happened when reading question "{q}" from "analysis values" sheet'.format(q=question_name),file=sys.stderr)

    def read_category_analysisvalue_column_historic(self,question_name,category_name):
        value = None
        category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
        if self.df_mdddata_categories is None or category_path not in self.df_mdddata_categories.index.get_level_values(0):
            raise self.CellNotFound('"label_mdd" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        # value = self.df_mdddata_categories.loc[category_path,sheet_mdddata_categories.column_names['col_analysisvalue']]
        # if not self.has_value_numeric(value):
        value = self.df_mdddata_categories.loc[category_path,sheet_mdddata_categories.column_names['col_analysisvalue_prev']]
        return value
        
    def read_category_analysisvalue_column_mdd(self,question_name,category_name):
        value = None
        category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
        if self.df_mdddata_categories is None or category_path not in self.df_mdddata_categories.index.get_level_values(0):
            raise self.CellNotFound('"label_mdd" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        value = self.df_mdddata_categories.loc[category_path,sheet_mdddata_categories.column_names['col_analysisvalue_mdd']]
        return value
        
    def read_category_analysisvalue(self,question_name,category_name):
        try:
            return self.read_category_analysisvalue_column_userinput(question_name,category_name)
        except self.CellNotFound:
            value = self.read_category_analysisvalue_column_historic(question_name,category_name)
            # if not self.has_value_text(value):
            #     value = self.read_category_analysisvalue_column_mdd(question_name,category_name)
            return value

    def read_category_label_column_userinput(self,question_name,category_name):
        def get_loc(question_name):
            # additional code to have better understanding of what happens when reading analysisvalues sheet
            # sorry, looks ugly a bit, 
            row_number = None
            _possible_pandas_error = None
            try:
                row_number = self.df_userinput_analysisvalues.index.get_loc(question_name)
            except Exception as e:
                _possible_pandas_error = e
            if _possible_pandas_error:
                if isinstance(_possible_pandas_error,KeyError):
                    raise self.__class__.ExcelMapException('ERROR: Reading "Analysis Values" sheet: can\'t find row index for "{f}" variable'.format(f=question_name))
                else:
                    raise _possible_pandas_error
            elif np.ndim(row_number)>0:
                raise self.__class__.ExcelMapException('ERROR: Reading "Analysis Values" sheet: Something is off, multiple entries for "{f}" variable'.format(f=question_name))
            else:
                pass # all good, continue
            return row_number
        try:
            value = None
            category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
            if self.df_userinput_analysisvalues is not None and question_name in self.df_userinput_analysisvalues.index.get_level_values(0):
                row_number = get_loc(question_name)
                assert np.isscalar(row_number), 'ERROR: something is off, from the excel with map: trying to find index of row "{row}" on "Analysis Values" sheet: row_number must be scalar numeric, got {var_type}'.format(var_type=type(row_number),row=question_name)
                row_category_names = self.df_userinput_analysisvalues.iloc[row_number,1:].squeeze().tolist()
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

                    row_category_labels = self.df_userinput_analysisvalues.iloc[row_number+1,1:].squeeze().tolist()
                    # row_category_analysisvalues = pd.Series(self.df_userinput_analysisvalues.iloc[row_number+2,1:],dtype=int).tolist()
                    # row_category_analysisvalues = pd.Series(self.df_userinput_analysisvalues.iloc[row_number+2,1:],dtype=int).astype(int).tolist()
                    # row_category_validation = self.df_userinput_analysisvalues.iloc[row_number+3,1:].squeeze().tolist()
                    index = row_category_names.index(category_name)
                    value = row_category_labels[index]
                    return value
            raise self.CellNotFound('"CAT_XXX" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=category_path,sheet_name=sheet_analysisvalues.sheet_name))
        except Exception as e:
            print('ERROR: something happened when reading labels for question == "{q}", cat == "{cat}" from "Analysis Values" sheet',file=sys.stderr)
            raise e

    def read_category_label_column_historic(self,question_name,category_name):
        value = None
        category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
        if self.df_mdddata_categories is None or category_path not in self.df_mdddata_categories.index.get_level_values(0):
            raise self.CellNotFound('"label_mdd" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        # value = self.df_mdddata_categories.loc[category_path,sheet_mdddata_categories.column_names['col_label']]
        # if not self.has_value_text(value):
        value = self.df_mdddata_categories.loc[category_path,sheet_mdddata_categories.column_names['col_label_prev']]
        return value
        
    def read_category_label_column_mdd(self,question_name,category_name):
        value = None
        category_path = '{var}.Categories[{cat}]'.format(var=question_name,cat=category_name)
        if self.df_mdddata_categories is None or category_path not in self.df_mdddata_categories.index.get_level_values(0):
            raise self.CellNotFound('"label_mdd" cell not found on "{sheet_name}" sheet for "{item}"'.format(item=question_name,sheet_name=sheet_mdddata_variables.sheet_name))
        value = self.df_mdddata_categories.loc[category_path,sheet_mdddata_categories.column_names['col_label_mdd']]
        return value
        
    def read_category_label(self,question_name,category_name):
        try:
            return self.read_category_label_column_userinput(question_name,category_name)
        except self.CellNotFound:
            value = self.read_category_label_column_historic(question_name,category_name)
            # if not self.has_value_text(value):
            #     value = self.read_category_label_column_mdd(question_name,category_name)
            return value

