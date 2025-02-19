

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










# this fn generates human readable description of what variable type is
# the expected result could be just "text" or "single-punch"
# or "single-punch grid" or "numeric grid"
# or "datetime, inside a loop, inside a block with fields..."
# or "text, a helper field of..."
# so it the logic sometimes looks complicated, it calls it recursively and conditionally depending on parent
# but this function is in fact not important, it only generates that text that is not used programmatically
# so if it's something too complex to understand and support, you can just replace this whole function with some more straightforward implementation
# so these complex codes can be safely deleted (I mean, all within compile_question_type_description())
def compile_question_type_description(mdmvar):
    def compile_type_description_this_level_field(mdmvar):
        if has_parent(mdmvar):
            count_sisters = 0
            mdmparent = get_parent(mdmvar)
            is_helperfield = mdmparent.ObjectTypeValue==0
            if not is_helperfield:
                for f in mdmparent.Fields:
                    if is_dataless(f):
                        continue
                    elif f.Name==mdmvar.Name:
                        continue
                    count_sisters = count_sisters + 1
            if is_iterative(mdmparent) and count_sisters == 0:
                if is_data_item(mdmvar):
                    if mdmvar.DataType==3 and mdmvar.MaxValue==1:
                        return 'single-punch grid'
                    elif mdmvar.DataType==3:
                        return 'multi-punch grid'
                    elif mdmvar.DataType==1 or mdmvar.DataType==6:
                        return 'numeric grid'
                    elif mdmvar.DataType==2:
                        return 'text grid'
                    elif mdmvar.DataType==5:
                        return 'datetime grid'
                    elif mdmvar.DataType==7:
                        return 'boolean grid'
        if mdmvar.ObjectTypeValue==27:
            return 'mdm document'
        elif mdmvar.ObjectTypeValue==1 or mdmvar.ObjectTypeValue==2:
            return 'loop'
        elif mdmvar.ObjectTypeValue==3:
            return 'block with fields'
        elif mdmvar.ObjectTypeValue==0:
            if mdmvar.DataType==0:
                return 'not a data variable (info node)'
            elif mdmvar.DataType==1 or mdmvar.DataType==6:
                return 'numeric'
            elif mdmvar.DataType==2:
                return 'text'
            elif mdmvar.DataType==3 and mdmvar.MaxValue==1:
                return 'single-punch'
            elif mdmvar.DataType==3:
                return 'multi-punch'
            elif mdmvar.DataType==4:
                return 'object'
            elif mdmvar.DataType==5:
                return 'datetime'
            elif mdmvar.DataType==7:
                return 'boolean flag'
            elif mdmvar.DataType==16:
                return 'ivariables instance'
            else:
                raise Exception('unrecognized DataType: {o}'.format(o=mdmvar.DataType))
        raise Exception('unrecognized ObjectTypeValue: {o}'.format(o=mdmvar.ObjectTypeValue))
    result = compile_type_description_this_level_field(mdmvar)
    mdmparent = get_parent(mdmvar) if has_parent(mdmvar) else None
    parent_part_processed = False
    if re.match(r'.*?\bgrid\b.*?',result):
        parent_part_processed = True
    if is_helper_field(mdmvar):
        parent_part_processed
        result = result + ', a helper field of var of type '+compile_question_type_description(mdmparent)
        return result
    while mdmparent:
        if not parent_part_processed:
            parent_part = compile_type_description_this_level_field(mdmparent)
            result = result + ', inside {article}{type}'.format(type=parent_part,article='a ' if re.match(r'^\s*?(?:loop|array|block).*?',parent_part) else '')
        mdmparent = get_parent(mdmparent) if has_parent(mdmparent) else None
        parent_part_processed = False
    return result















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
            # for f in list_mdmcategories(mdmcat):
            #     yield f
            yield from list_mdmcategories(mdmcat)

def list_mdmcategories_with_slname(mdmvariable,skip_shared_lists=False):
    for mdmcat in mdmvariable.Elements:
        if mdmcat.IsReference:
            try:
                sl_name_clean = re.sub(r'[\^\\/\.]','',mdmcat.ReferenceName,flags=re.I|re.DOTALL)
                mdmsharedlist = mdmvariable.Document.Types[sl_name_clean]
                for _, mdmcat in list_mdmcategories_with_slname(mdmsharedlist):
                    yield mdmsharedlist.Name, mdmcat # we return outer so that we are checking that values are unique within outer most list
            except Exception as e:
                raise Exception('Was not able to refer to a Shared List "{l}": {e}'.format(l=mdmcat.ReferenceName,e=e)) from e
        elif mdmcat.Type==0:
            yield None, mdmcat
        else:
            # for sl_name, mdmcat in list_mdmcategories_with_slname(mdmcat):
            #     yield sl_name, mdmcat
            yield from list_mdmcategories_with_slname(mdmcat)



def list_mdmdatafields_recursively(mdmvariable):
    for mdmfield in ( [f for f in mdmvariable.Fields] if has_mdmfield_nested_fields(mdmvariable) else [] ) + [f for f in mdmvariable.HelperFields]:
        if is_data_item(mdmfield):
            yield mdmfield
        yield from list_mdmdatafields_recursively(mdmfield)
    yield from []

