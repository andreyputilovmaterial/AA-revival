

import re


def has_mdmfield_nested_fields(mdmvariable):
    return mdmvariable.ObjectTypeValue==1 or mdmvariable.ObjectTypeValue==2 or mdmvariable.ObjectTypeValue==3



def is_data_item(mdmvariable):
    return mdmvariable.ObjectTypeValue==0 and not (mdmvariable.DataType==0)



# def is_loop(mdmvariable):
#     return mdmvariable.ObjectTypeValue==1 or mdmvariable.ObjectTypeValue==2



def is_nocasedata(mdmvariable):
    return mdmvariable.HasCaseData==False



def is_dataless(mdmvariable):
    return mdmvariable.ObjectTypeValue==0 and mdmvariable.DataType==0



def is_system(mdmvariable):
    return mdmvariable.IsSystem==True



def has_own_categories(mdmvariable):
    return is_categorical(mdmvariable)



def is_iterative(mdmvariable):
    return mdmvariable.ObjectTypeValue==1 or mdmvariable.ObjectTypeValue==2



def is_categorical(mdmvariable):
    return mdmvariable.ObjectTypeValue==0 and mdmvariable.DataType==3



def has_parent(mdmvar):
    return mdmvar.Parent and mdmvar.Parent.Name



def get_parent(mdmvar):
    if mdmvar.Parent and mdmvar.Parent.Name:
        mdmparent = mdmvar.Parent
        if '@class' in mdmparent.Name:
            return get_parent(mdmparent)
        else:
            return mdmparent
    return None



def is_helper_field(mdmvar):
    if not has_parent(mdmvar):
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



def list_mdmcategories(mdmvariable,skip_shared_lists=False):
    for mdmcat in mdmvariable.Elements:
        if mdmcat.IsReference:
            try:
                sl_name_clean = re.sub(r'[\^\\/\.]','',mdmcat.ReferenceName,flags=re.I|re.DOTALL)
                mdmsharedlist = mdmvariable.Document.Types[sl_name_clean]
                for f in list_mdmcategories(mdmsharedlist):
                    yield f
            except Exception as e:
                raise Exception('Was not able to refer to a Shared List "{l}": {e}'.format(l=mdmcat.ReferenceName,e=e)) from e
        elif mdmcat.Type==0:
            yield mdmcat
        else:
            for f in list_mdmcategories(mdmcat):
                yield f




