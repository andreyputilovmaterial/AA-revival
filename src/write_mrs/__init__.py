import re








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



def make_path_to_field(variable,MDD):
    if MDD.is_helper_field(variable):
        return make_path_to_field(variable.Parent,MDD) + '.HelperFields["{name}"]'.format(name=variable.Name)
    if '@class' in variable.Name:
        return make_path_to_field(variable.Parent,MDD)
    # path = re.replace(r'\[[^\]]*?\]','',variable.FullName,flags=re.I|re.DOTALL)
    # parts = path.split('.')
    # return 'mdm' + ''.join(['.Fields["{name}"]'.format(name=part) for part in parts])
    if MDD.has_parent(variable):
        return make_path_to_field(variable.Parent,MDD) + '.Fields["{name}"]'.format(name=variable.Name)
    else:
        return 'mDocument' + '.Fields["{name}"]'.format(name=variable.Name)






def generate_savprep_mrs_include_part(mdd,map,config):
    assert not not mdd, 'Error: MDD is required but it missing'
    mddvariables = mdd.variables

    result = ''

    result = result + '\n'
    result = result + '\' AnalysisAuthorRevival\n'
    result = result + '\' Generated: {t} (time configured as system time on a local machine where the script was running, without timezone information)\n'.format(t=config['datetime'])
    result = result + '\n'

    for variable in mddvariables:
        if variable.Name=='':
            continue
        question_name = variable.FullName

        is_included            = map.read_question_is_included(question_name)
        is_data_variable       = mdd.is_data_item(variable) # map.read_question_...()
        is_iterative           = mdd.is_iterative(variable) or mdd.has_own_categories(variable)
        question_shortname     = map.read_question_shortname(question_name)
        question_label         = map.read_question_label(question_name)

        result = result + '\n'
        result = result + '\' {name}\n'.format(name=question_name)
        result = result + 'debug.echo("processing {name}...")\n'.format(name=question_name)
        result = result + 'debug.log("processing {name}...")\n'.format(name=question_name)
        result = result + 'set oField = {address_field}\n'.format(address_field=make_path_to_field(variable,mdd))

        if is_iterative:
            for category in mdd.list_mdmcategories(variable):
                category_name = category.Name

                category_analysisvalue = map.read_category_analysisvalue(question_name,category_name)
                if isinstance(category_analysisvalue,str):
                    if re.match(r'^\s*?$',category_analysisvalue,flags=re.I|re.DOTALL): # if empty, set true emptyness
                        category_analysisvalue = None
                category_label = map.read_category_label(question_name,category_name)

                if is_included:
                    if category_analysisvalue is None:
                        raise Exception('Error: category analysis value is missing for "{cat}"'.format(cat='{var}.Categories["{catname}"]'.format(var=question_name,catname=category_name)))

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
