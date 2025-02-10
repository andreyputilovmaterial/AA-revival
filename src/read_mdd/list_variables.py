
import re


if __name__ == '__main__':
    # run as a program
    import load_mdd
elif '.' in __name__:
    # package
    from . import load_mdd
else:
    # included with no parent package
    import load_mdd





def has_mdmfield_nested_fields(mdmvariable):
    return mdmvariable.ObjectTypeValue==1 or mdmvariable.ObjectTypeValue==2 or mdmvariable.ObjectTypeValue==3




def list_mdmfields(mdmvariable):
    for mdmitem in mdmvariable:
        yield mdmitem
        # if array, or grid, or block (it means, if it has subfields)
        if has_mdmfield_nested_fields(mdmitem):
            for f in list_mdmfields(mdmitem.Fields):
                yield f
        # and process HelperFields
        for f in list_mdmfields(mdmitem.HelperFields):
            yield f



# def list_mdmcategories(mdmvariable,skip_shared_lists=False):
#     for mdmcat in mdmvariable.Elements:
#         if mdmcat.IsReference:
#             try:
#                 sl_name_clean = re.sub(r'[\^\\/\.]','',mdmcat.ReferenceName,flags=re.I|re.DOTALL)
#                 mdmsharedlist = mdmvariable.Document.Types[sl_name_clean]
#                 for f in list_mdmcategories(mdmsharedlist):
#                     yield f
#             except Exception as e:
#                 raise Exception('Was not able to refer to a Shared List "{l}": {e}'.format(l=mdmcat.ReferenceName,e=e)) from e
#         elif mdmcat.Type==0:
#             yield mdmcat
#         else:
#             for f in list_mdmcategories(mdmcat):
#                 yield f




def get_variables(mdd_file_name):
    mdmroot = load_mdd.open_mdd_document(mdd_file_name)
    mdmvariables = [ mdmroot ] + [ mdmitem for mdmitem in list_mdmfields(mdmroot.Fields) ]

    return mdmvariables
