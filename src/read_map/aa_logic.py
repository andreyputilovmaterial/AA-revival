


import re


from . import util_mdmvars



def should_exclude(mdmitem):
    def should_exclude_based_on_props(mdmitem):
        return  mdmitem.Properties['SavRemove'] or mdmitem.Properties['D_RemoveSav'] or mdmitem.Properties['D_Remove'] or mdmitem.Properties['RemoveSav']
    def should_exclude_recursively(mdmitem):
        if should_exclude_based_on_props(mdmitem):
            return True
        mdmparent = util_mdmvars.get_parent(mdmitem) if util_mdmvars.has_parent(mdmitem) else None
        if mdmparent:
            if should_exclude_recursively(mdmparent):
                return True
        return False
    if util_mdmvars.is_nocasedata(mdmitem):
        return True
    return should_exclude_recursively(mdmitem)



def sanitize_shortname(name):
    def append3decimalplaces(name):
        return '000'[:3-len(name)]+name
    if name==0:
        return append3decimalplaces('0')
    elif not name:
        return ''
    else:
        name = '{s}'.format(s=name)
    name = re.sub(r'^\s*?(\d+)(?:\.\d*)?\s*?$',lambda m: append3decimalplaces(m[1]),name,flags=re.I|re.DOTALL)
    return name



def sanitize_analysis_value(value):
    if value==0:
        return 0
    if not value:
        return None
    value = '{v}'.format(v=value)
    value = re.sub(r'^\s*?(\d+)(?:\.0\d*)\s*?$',lambda m: m[1],value,flags=re.I|re.DOTALL)
    try:
        value = float(value)
        try:
            value_rounded = int(round(value))
            if abs(value_rounded-value)<0.001:
                value = value_rounded
        except:
            pass
    except:
        pass
    return value



def validate_shortname(name):
    return not not re.match(r'^\s*?(?:\w+|(?:\[L:[^\]:]*:[^:\]]*\]))+\s*?$',name,flags=re.I|re.DOTALL) and not re.match(r'^\s*?\d.*?$',name,flags=re.I|re.DOTALL)



def read_shortname(mdmvar):
    shortname = None
    try:
        shortname = read_shortname_aastyle_logic_based_on_properties(mdmvar)
    except AAFailedFindShortnameException:
        if util_mdmvars.has_parent(mdmvar) and util_mdmvars.is_helper_field(mdmvar):
            try:
                shortname = read_shortname_aastyle_otherspec(mdmvar)
            except AAFailedFindShortnameException:
                shortname = read_shortname_fallback(mdmvar)
        else:
            shortname = read_shortname_fallback(mdmvar)
    # assert validate_shortname(shortname), 'ShortName does not seem to be valid, please check'
    if not validate_shortname(shortname):
        # a place for breakpoint
        # raise AssertionError('ShortName does not seem to be valid, please check: "{shortname}"'.format(shortname=shortname))
        pass
    return shortname



class AAFailedFindShortnameException(Exception):
    """AAFailedFindShortnameException"""
def read_shortname_aastyle_logic_based_on_properties(mdmvar):
    try:
        if not util_mdmvars.has_parent(mdmvar):
            result = sanitize_shortname(mdmvar.Properties['ShortName'])
            if not result:
                raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))
            return result
        else:
            count_sisters = 0
            mdmparent = util_mdmvars.get_parent(mdmvar)
            is_numeric_grid = None
            is_text_grid = None # I am treating None and False differently here, which is not best design probably
            is_categorical_grid = None # I am treating None and False differently here, which is not best design probably
            is_helper_field = util_mdmvars.is_helper_field(mdmvar)
            if not is_helper_field:
                for f in mdmparent.Fields:
                    if util_mdmvars.is_dataless(f):
                        continue
                    elif f.Name==mdmvar.Name:
                        continue
                    count_sisters = count_sisters + 1
            if util_mdmvars.is_iterative(mdmparent):
                if util_mdmvars.is_data_item(mdmvar):
                    if mdmvar.DataType==3:
                        if count_sisters==0:
                            if mdmvar.MaxValue==1:
                                # 'single-punch grid'
                                is_categorical_grid = True
                            else:
                                # 'multi-punch grid'
                                is_categorical_grid = True
                    if mdmvar.DataType==1 or mdmvar.DataType==6:
                        # 'numeric grid'
                        # if all found fields are numeric (DataType==1/long or DataType==6/float)
                        # so we set it to True if it was unset (if it's still None)
                        # but we don't do if it's False, it means, a non-numeric field was found
                        if is_numeric_grid is None:
                            is_numeric_grid = True
                    elif is_numeric_grid:
                        # if it was set to True but the field is not numeric - we set it to False,
                        # and if we find a numeric field again, we don't set it to True again, because we only do if it's None and we don't do if it's False
                        # so I am treating None and False differently here, which is not best design probably
                        is_numeric_grid = False
                    if mdmvar.DataType==2:
                        # text grid, similar to above
                        if is_text_grid is None:
                            is_text_grid = True
                    elif is_text_grid:
                        is_text_grid = False
                    if mdmvar.DataType==5:
                        pass # 'datetime grid' - we don't capture this type of grids
                    elif mdmvar.DataType==7:
                        pass # 'boolean grid' - we don't capture this type of grids
            

            if is_numeric_grid or is_text_grid:
                result_part1 = sanitize_shortname(mdmparent.Properties['ShortName'])
                result_part2 = sanitize_shortname(mdmvar.Properties['ShortName'])
                if not result_part1:
                    raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))
                if not result_part2:
                    raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))
                return '{p1}{add}{p2}'.format(p1=result_part1,p2=result_part2,add='_') # TODO: check, maybe iterative part should be inserted in the middle
            elif is_categorical_grid:
                result = sanitize_shortname(mdmvar.Properties['ShortName'])
                if not result:
                    raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))
                return result
            else:
                result = sanitize_shortname(mdmvar.Properties['ShortName'])
                if not result:
                    raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))
                return result
    except AAFailedFindShortnameException as e:
        if util_mdmvars.has_parent(mdmvar) and util_mdmvars.is_helper_field(mdmvar):
            result_part1 = mdmparent.Properties["ShortName"]
            result_part2 = mdmvar.Properties["ShortName"]
            if not result_part1:
                raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))
            if not result_part2:
                raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))
            return '{p1}{add}{p2}'.format(p1=result_part1,p2=result_part2,add='_')
        else:
            raise e




# for Ethnicity.HelperFields[AnotherRace]
# (when the ShortName for Ethnicyt is S18, and analysis value for "AnotherRace" is 98)
# AA was generating
# S18_Other_98
# we are trying to replicate the same logic here
# we don't have our map passed to a function, so we can't grab the code [98] here
# so I return [L:Ethnicity:AnotherRace]
# and then, later, it is replaced with an Excel function
def read_shortname_aastyle_otherspec(mdmvar):
    def is_equal_name(name1,name2):
        return '{n}'.format(n=name1).lower()=='{n}'.format(n=name2).lower()
    mdmcat_matching_otherspec = None
    mdmparent = util_mdmvars.get_parent(mdmvar) if util_mdmvars.has_parent(mdmvar) else None
    if mdmparent and util_mdmvars.is_helper_field(mdmvar):
        # trying to find that "Other" category that matches that given field
        if util_mdmvars.is_categorical(mdmparent):
            for mdmcat in util_mdmvars.list_mdmcategories(mdmparent):
                if mdmcat.Flag==80 and mdmcat.OtherReference and is_equal_name(mdmcat.Name,mdmvar.Name):
                    mdmcat_matching_otherspec = mdmcat
    if mdmcat_matching_otherspec:
        # previously I had
        # result = '="'+shortname_parent_part + '_Other_' +'" & VLOOKUP("'+mdmcat.Name+'"\'MDD_Data_Categories\'!$
        # this is super bad design, returning a formula from here
        # if the layout of columns changes, no one will find out that it needs to be updated here
        # we are expected to return a name (a string value) from this function
        # but I am trying to repeat older logic from AA, we need category code here, I can't think of better solution right now
        # maybe I should introduce some syntax with inserts, like [L0], or [L0:category], something following similar standards to flatout pattern
        # not sure, maybe I'll do
        shortname_parent_part = read_shortname_aastyle_logic_based_on_properties(mdmparent)
        # return Parent+'_Other_[L0]'
        # result = '="'+shortname_parent_part + '_Other_' +'" & VLOOKUP("'+mdmcat.Name+'"\'MDD_Data_Categories\'!$A$2:$A$999999,6.FALSE)'
        result = shortname_parent_part + '_Other_' +'[L:{var}:{cat}]'.format(var=mdmparent.FullName,cat=mdmcat_matching_otherspec.Name)
        return result
    else:
        raise AAFailedFindShortnameException('Failed to find shortname: {s}'.format(s=mdmvar.Name))



def read_shortname_fallback(mdmvar):
    def check_if_improper_name(name):
        is_improper_name = False
        # and there are less common cases but still happening in disney bes
        is_improper_name = is_improper_name or not not re.match(r'^\s*?(\d+)\s*?$',name,flags=re.I)
        is_improper_name = is_improper_name or not not re.match(r'^\s*?((?:Top|T|Bottom|B))(\d*)((?:B|Box))\s*?$',name,flags=re.I)
        is_improper_name = is_improper_name or not not re.match(r'^\s*?(?:GV|Rank|Num)\s*?$',name,flags=re.I)
        return is_improper_name
    if mdmvar.Name=='':
        return None
    field_prop_shortname = sanitize_shortname(mdmvar.Properties['ShortName'])
    if not field_prop_shortname:
        field_prop_shortname = mdmvar.Name
    
    sisters = []
    first_field = None
    mdmparent = util_mdmvars.get_parent(mdmvar)

    is_impropert_field_own_name = False
    is_impropert_field_own_name = is_impropert_field_own_name or not field_prop_shortname
    is_impropert_field_own_name = is_impropert_field_own_name or check_if_improper_name(field_prop_shortname)
    already_used = [] # TODO: 
    is_impropert_field_own_name = is_impropert_field_own_name or already_used

    if not util_mdmvars.has_parent(mdmvar):
        if is_impropert_field_own_name:
            return mdmvar.Name
        else:
            return field_prop_shortname
    else:
        if util_mdmvars.is_helper_field(mdmvar):
            return read_shortname(mdmparent) + '_' + ( field_prop_shortname if field_prop_shortname else mdmvar.Name )
        else:
            for f in mdmparent.Fields:
                if not first_field:
                    first_field = f
                if util_mdmvars.is_dataless(f):
                    continue
                elif f.Name==mdmvar.Name:
                    continue
                sisters.append(f)
        
            if is_impropert_field_own_name:
                if len(sisters)==0:
                    return read_shortname(mdmparent)
                elif check_if_improper_name(mdmvar.Name) and first_field and mdmvar.Name==first_field.Name:
                    return read_shortname(mdmparent)
                else:
                    return read_shortname(mdmparent) + '_' + ( field_prop_shortname if field_prop_shortname else mdmvar.Name )
            else:
                return field_prop_shortname








