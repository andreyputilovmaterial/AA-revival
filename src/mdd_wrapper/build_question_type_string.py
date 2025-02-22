

import re



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
def compile_question_type_description(mdmvar,MDM):
    def compile_type_description_this_level_field(mdmvar):
        if MDM.has_parent(mdmvar):
            count_sisters = 0
            mdmparent = MDM.get_parent(mdmvar)
            is_helperfield = mdmparent.ObjectTypeValue==0
            if not is_helperfield:
                for f in mdmparent.Fields:
                    if MDM.is_dataless(f):
                        continue
                    elif f.Name==mdmvar.Name:
                        continue
                    count_sisters = count_sisters + 1
            if MDM.is_iterative(mdmparent) and count_sisters == 0:
                if MDM.is_data_item(mdmvar):
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
            return 'MDM document'
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
            elif mdmvar.DataType==27:
                return 'MDM document'
            else:
                raise Exception('unrecognized DataType: {o}'.format(o=mdmvar.DataType))
        raise Exception('unrecognized ObjectTypeValue: {o}'.format(o=mdmvar.ObjectTypeValue))
    result = compile_type_description_this_level_field(mdmvar)
    mdmparent = MDM.get_parent(mdmvar) if MDM.has_parent(mdmvar) else None
    parent_part_processed = False
    if re.match(r'.*?\bgrid\b.*?',result):
        parent_part_processed = True
    if MDM.is_helper_field(mdmvar):
        parent_part_processed
        result = result + ', a helper field of var of type '+compile_question_type_description(mdmparent,MDM)
        return result
    while mdmparent:
        if not parent_part_processed:
            parent_part = compile_type_description_this_level_field(mdmparent)
            result = result + ', inside {article}{type}'.format(type=parent_part,article='a ' if re.match(r'^\s*?(?:loop|array|block).*?',parent_part) else '')
        mdmparent = MDM.get_parent(mdmparent) if MDM.has_parent(mdmparent) else None
        parent_part_processed = False
    return result
