
from pathlib import Path
import re




import_er = None
try:
    import win32com.client
except Exception as e:
    import_er = e



from . import build_question_type_string as question_type


class ReadMDDError(Exception):
    """Read MDD error"""




class MDD:
    def __init__(self,mdd_path):
        try:
            if import_er:
                raise import_er # we only fail here, we don't fail if we don't need to read MDD
            mdmDocument = win32com.client.Dispatch("MDM.Document")
            # openConstants_oNOSAVE = 3
            openConstants_oREAD = 1
            # openConstants_oREADWRITE = 2
            # print('Oopening MDM document using method "open": "{path}"'.format(path=mdd_path))
            # we'll check that the file exists so that the error message is more informative - otherwise you see a long stack of messages that do not tell much
            if not(Path(mdd_path).is_file()):
                raise FileNotFoundError('File not found: {fname}'.format(fname=mdd_path))
            mdmDocument.Open( mdd_path, "", openConstants_oREAD )
            self.mdmDocument = mdmDocument
        except Exception as e:
            raise ReadMDDError('Failed to open MDD: {e}'.format(e=e)) from e
        
    



    @property
    def variables(self):
        # return [ self.mdmDocument ] + [ mdmitem for mdmitem in MDD.list_mdmfields(self.mdmDocument.Fields) ]
        return [ mdmitem for mdmitem in MDD.list_mdmfields(self.mdmDocument.Fields) ] # we need to know progress, so we need len(), that's why I return list not generator
        # return  MDD.list_mdmfields(self.mdmDocument.Fields) # I'll return a generator, not a list







    @staticmethod
    def has_mdmfield_nested_fields(mdmvariable):
        return mdmvariable.ObjectTypeValue==1 or mdmvariable.ObjectTypeValue==2 or mdmvariable.ObjectTypeValue==3

    @staticmethod
    def is_data_item(mdmvariable):
        return mdmvariable.ObjectTypeValue==0 and not (mdmvariable.DataType==0)

    # def is_loop(mdmvariable):
    #     return mdmvariable.ObjectTypeValue==1 or mdmvariable.ObjectTypeValue==2

    @staticmethod
    def is_nocasedata(mdmvariable):
        return mdmvariable.HasCaseData==False

    @staticmethod
    def is_dataless(mdmvariable):
        return mdmvariable.ObjectTypeValue==0 and mdmvariable.DataType==0

    @staticmethod
    def is_system(mdmvariable):
        return mdmvariable.IsSystem==True

    @staticmethod
    def has_own_categories(mdmvariable):
        return MDD.is_categorical(mdmvariable)

    @staticmethod
    def is_iterative(mdmvariable):
        return mdmvariable.ObjectTypeValue==1 or mdmvariable.ObjectTypeValue==2

    @staticmethod
    def is_categorical(mdmvariable):
        return mdmvariable.ObjectTypeValue==0 and mdmvariable.DataType==3

    @staticmethod
    def has_parent(mdmvar):
        return mdmvar.Parent and mdmvar.Parent.Name

    # get_parent()
    # it's better to always use this function
    # so don't call "parent = item.Parent"
    # please do "parent = get_parent(item)" instead
    # because, "item.Parent" can give you a meaningless block of fields with no name
    # and this funciton skips it and gives the next parent
    # for example
    # the metadata scripts are
    # Familiarity "How familiar are you..." loop { use SL_BrandList } fields - ( GV "" categorical...
    # and your "item" variable references to "GV"
    # you check item.Parent
    # and you are not getting "Familiarity"
    # you are getting some middle name that stands for "fields" which is meaningless, and you need another step to get to "Familiarity"
    # so checking for grids is getting more complicated
    @staticmethod
    def get_parent(mdmvar):
        if mdmvar.Parent and mdmvar.Parent.Name:
            mdmparent = mdmvar.Parent
            if '@class' in mdmparent.Name:
                return MDD.get_parent(mdmparent)
            else:
                return mdmparent
        return None

    @staticmethod
    def is_helper_field(mdmvar):
        if not MDD.has_parent(mdmvar):
            return False
        mdmparent = mdmvar.Parent
        if mdmparent.ObjectTypeValue==0:
            return True
        if mdmparent.HelperFields.Exist(mdmvar.Name):
            return True
        else:
            return False

    # def is_within_loop(mdmvar):
    #     return 

    # compile_question_type_description()
    # this fn generates human readable description of what variable type is
    # the expected result could be just "text" or "single-punch"
    # or "single-punch grid" or "numeric grid"
    # or "datetime, inside a loop, inside a block with fields..."
    # or "text, a helper field of..."
    # so it the logic sometimes looks complicated, it calls it recursively and conditionally depending on parent
    # but this function is in fact not important, it only generates that text that is not used programmatically
    # so if it's something too complex to understand and support, you can just replace this whole function with some more straightforward implementation
    # so these complex codes can be safely deleted (I mean, all within compile_question_type_description())
    @staticmethod
    def compile_question_type_description(mdmvar):
        return question_type.compile_question_type_description(mdmvar,MDD)

    @staticmethod
    def list_mdmfields(mdmvariable):
        for mdmitem in mdmvariable:
            yield mdmitem
            # if array, or grid, or block (it means, if it has subfields)
            if MDD.has_mdmfield_nested_fields(mdmitem):
                yield from MDD.list_mdmfields(mdmitem.Fields)
            # and process HelperFields
            yield from MDD.list_mdmfields(mdmitem.HelperFields)

    @staticmethod
    def list_mdmcategories(mdmvariable,skip_shared_lists=False):
        for mdmcat in mdmvariable.Elements:
            if mdmcat.IsReference:
                # first, let's handle shared list reference
                try:
                    sl_name_clean = re.sub(r'[\^\\/\.]','',mdmcat.ReferenceName,flags=re.I|re.DOTALL)
                    mdmsharedlist = mdmvariable.Document.Types[sl_name_clean]
                    yield from MDD.list_mdmcategories(mdmsharedlist)
                except Exception as e:
                    raise Exception('Was not able to refer to a Shared List "{l}": {e}'.format(l=mdmcat.ReferenceName,e=e)) from e
            elif mdmcat.Type==0:
                # simple category
                yield mdmcat
            elif mdmcat.Type==13:
                # sublist - iterate over everything in the inside
                # for f in list_mdmcategories(mdmcat):
                #     yield f
                yield from MDD.list_mdmcategories(mdmcat)
            elif mdmcat.Type==2:
                # analysisbase element - do not yield anything, skip - not a data item
                pass
            else:
                # every we did not handle before
                try:
                    # for f in list_mdmcategories(mdmcat):
                    #     yield f
                    yield from MDD.list_mdmcategories(mdmcat)
                except Exception as e:
                    print('Debugging: failed when processing category {mdmcat}, mdmcat.Type == {t}, Error: {e}'.format(mdmcat=mdmcat.Name,t=mdmcat.Type,e=e))
                    yield mdmcat

    @staticmethod
    def list_mdmcategories_with_slname(mdmvariable,skip_shared_lists=False):
        for mdmcat in mdmvariable.Elements:
            if mdmcat.IsReference:
                # first, let's handle if it's a reference to a shared list
                try:
                    sl_name_clean = re.sub(r'[\^\\/\.]','',mdmcat.ReferenceName,flags=re.I|re.DOTALL)
                    mdmsharedlist = mdmvariable.Document.Types[sl_name_clean]
                    for _, mdmcat in MDD.list_mdmcategories_with_slname(mdmsharedlist):
                        yield mdmsharedlist.Name, mdmcat # we return outer so that we are checking that values are unique within outer most list
                except Exception as e:
                    raise Exception('Was not able to refer to a Shared List "{l}": {e}'.format(l=mdmcat.ReferenceName,e=e)) from e
            elif mdmcat.Type==0:
                # simple plain category - just return it
                yield None, mdmcat
            elif mdmcat.Type==13:
                # sublist - iterate over everything in the inside
                yield from MDD.list_mdmcategories_with_slname(mdmcat)
            elif mdmcat.Type==2:
                # analysisbase element - no need to print, not a data item
                pass
            else:
                # something we could not handle before - we'll try to iterate over inner items
                # for sl_name, mdmcat in list_mdmcategories_with_slname(mdmcat):
                #     yield sl_name, mdmcat
                try:
                    yield from MDD.list_mdmcategories_with_slname(mdmcat)
                except Exception as e:
                    print('For debugging: processing "{mdmcat}", mdmcat.Type=={t}: can\'t process'.format(mdmcat=mdmcat.Name,t=mdmcat.Type))
                    yield None, mdmcat
                    # raise e

    @staticmethod
    def list_mdmdatafields_recursively(mdmvariable):
        for mdmfield in ( [f for f in mdmvariable.Fields] if MDD.has_mdmfield_nested_fields(mdmvariable) else [] ) + [f for f in mdmvariable.HelperFields]:
            if MDD.is_data_item(mdmfield):
                yield mdmfield
            yield from MDD.list_mdmdatafields_recursively(mdmfield)
        yield from []
