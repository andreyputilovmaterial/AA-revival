


from . import load as load_m
from . import update_map as update_map_m
from . import write_to_file as write_to_file_m



def load(*args,**kwargs):
    return load_m.load(*args,**kwargs)

def update_map(*args,**kwargs):
    return update_map_m.update_map(*args,**kwargs)

def write_to_file(*args,**kwargs):
    return write_to_file_m.write_to_file(*args,**kwargs)
