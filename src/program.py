from datetime import datetime, timezone
import argparse
from pathlib import Path
import traceback
import shutil



import codecs



if __name__ == '__main__':
    # run as a program
    import write_mrs
    from mdd_wrapper import MDD
    from map_wrapper import Map
elif '.' in __name__:
    # package
    from . import write_mrs
    from .mdd_wrapper import MDD
    from .map_wrapper import Map
else:
    # included with no parent package
    import write_mrs
    from mdd_wrapper import MDD
    from map_wrapper import Map




def create_backup(out_filename,config):
    script_run_datetime = config['datetime']
    script_run_datetime = script_run_datetime.strftime('_%Y%m%d_%I%M%S_Z')
    out_bkp_filename = out_filename.with_stem(out_filename.stem+'_backup_'+script_run_datetime)
    from_file = Path(out_filename)
    to_file = Path(out_bkp_filename)
    shutil.copy(from_file,to_file)






def cli_program_update_map():
    time_start = datetime.now()
    script_name = 'AA Revival: update map'

    parser = argparse.ArgumentParser(
        description="mdmtools AA Revival utility: update map",
        prog='mddtools-aarevival-update-map'
    )
    parser.add_argument(
        '--mdd',
        type=str,
        help='Path to MDD',
        required=False
    )
    parser.add_argument(
        '--map',
        help='Map file name',
        type=str,
        required=False
    )
    parser.add_argument(
        '--config-fill-analysisvalues',
        help='Config option how to handle analysis values: load from MDD, pre-fill, leave blank',
        choices=['from_mdd_or_blank','from_mdd_and_autofill','autofill_override','blank'],
        type=str,
        required=False
    )
    parser.add_argument(
        '--create-map-if-not-exists',
        help='Config option indicating if a map should be created if a linked map file is not found',
        type=str,
        required=False
    )
    parser.add_argument( '--no-map-backup', help='Suppress creating a backup?', type=str, required=False, )
    args, _ = parser.parse_known_args()
    
    config = {}
    config['datetime'] = time_start
    config['skip_map_backup'] = False
    if args.no_map_backup:
        config['skip_map_backup'] = True
    config['create_map_if_not_exists'] = False
    if args.create_map_if_not_exists:
        config['create_map_if_not_exists'] = True

    if args.config_fill_analysisvalues:
        config['map_options'] = {
            'process_analysis_values': args.config_fill_analysisvalues,
        }
    
    mdd_filename = None
    mdd = None
    if args.mdd:
        mdd_filename = Path(args.mdd)
        mdd_filename = '{fname}'.format(fname=mdd_filename.resolve())
        if not(Path(mdd_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=mdd_filename))
        print('{script_name}: reading mdd "{fname}"'.format(fname=mdd_filename,script_name=script_name))
        mdd = MDD(mdd_filename)

    map_filename = None
    map_df = None
    if args.map:
        map_filename = Path(args.map)
        map_filename = '{fname}'.format(fname=map_filename.resolve())
        if Path(map_filename).is_file():
            print('{script_name}: reading map "{fname}"'.format(fname=map_filename,script_name=script_name))
            map_df = Map(map_filename,config)
        else:
            if 'create_map_if_not_exists' in config and config['create_map_if_not_exists']:
                pass
            else:
                raise FileNotFoundError('file not found: {fname}'.format(fname=map_filename))
    
    config['mdd_file_provided'] = not not mdd_filename
    config['map_file_provided'] = not not map_filename
    config['mdd_file_exists'] = not not map_df
    config['map_file_exists'] = not not mdd
    config['mdd_filename'] = mdd_filename if mdd_filename else ''
    config['map_filename'] = map_filename if map_filename else ''

    if not mdd_filename:
        if not map_filename:
            raise FileNotFoundError('MDD not map file are provided. You have to pass MDD file name and/or map file name. If MDD file name is not provided, it can be read from the map. If map file is not provided, it will be created with a default name. But you have to set of of those, --mdd or --map, or both.')
        print('{script_name}: MDD name not provided, reading MDD path from the map'.format(script_name=script_name))
        mdd_filename = map.get_mdd_path()
        mdd_filename = Path(mdd_filename)
        if not mdd_filename.is_absolute():
            # relative to AA Excel Map file
            mdd_filename = ( ( Path(map_filename) if Path(map_filename).is_dir() else Path(map_filename).parents[0] ) if map_filename else Path('.') ) / mdd_filename
        mdd_filename = '{fname}'.format(fname=mdd_filename.resolve())
        if not(Path(mdd_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=mdd_filename))
        print('{script_name}: reading MDD "{fname}"'.format(fname=mdd_filename,script_name=script_name))
        mdd = MDD(mdd_filename)
        config['mdd_filename'] = mdd_filename if mdd_filename else ''

    out_filename = map_filename if map_filename else ( ( Path(mdd_filename) if Path(mdd_filename).is_dir() else Path(mdd_filename).parents[0] ) if mdd_filename else Path('.') ) / 'AnalysisAuthorRevival.xlsx'
    out_filename = Path(out_filename)
    if Path(map_filename).is_file():
        if not config['skip_map_backup']:
            print('{script_name}: creating backup of the map: {map}'.format(script_name=script_name,map=out_filename))
            create_backup(out_filename,config)
    assert out_filename, 'out filename is still missing, please check the code'

    print('{script_name}: working, updating the map'.format(script_name=script_name))
    result_df = map_df.update(mdd)
    
    print('{script_name}: saving as "{fname}"'.format(fname=out_filename,script_name=script_name))
    result_df.write_to_file(out_filename)

    time_finish = datetime.now()
    print('{script_name}: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start,script_name=script_name))




def cli_program_produce_savprep_mrs():
    time_start = datetime.now()
    script_name = 'AA Revival: produce savprep mrs'

    parser = argparse.ArgumentParser(
        description="mdmtools AA Revival utility: produce savprep mrs",
        prog='mddtools-aarevival-produce-savprep-mrs'
    )
    parser.add_argument(
        '--mdd',
        type=str,
        help='Path to MDD',
        required=True
    )
    parser.add_argument(
        '--map',
        help='Map file name',
        type=str,
        required=True
    )
    parser.add_argument(
        '--out',
        help='Output file name',
        type=str,
        required=False
    )
    # parser.add_argument( '--write-template', help='Do we need to create a 601_SavPrepRevival.mrs?', type=str, required=False, )
    args, _ = parser.parse_known_args()
    
    config = {}
    # config['need_write_template'] = False
    # if args.write_template:
    #     config['need_write_template'] = True
    config['datetime'] = time_start

    mdd = None
    if args.mdd:
        mdd_filename = Path(args.mdd)
        mdd_filename = '{fname}'.format(fname=mdd_filename.resolve())
        if not(Path(mdd_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=mdd_filename))
        print('{script_name}: reading mdd "{fname}"'.format(fname=mdd_filename,script_name=script_name))
        mdd = MDD(mdd_filename)
    else:
        raise FileNotFoundError('MDD: file not provided; please use --mdd option')

    map_df = None
    if args.map:
        map_filename = Path(args.map)
        map_filename = '{fname}'.format(fname=map_filename.resolve())
        if not(Path(map_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=map_filename))
        print('{script_name}: reading map "{fname}"'.format(fname=map_filename,script_name=script_name))
        map_df = Map(map_filename,config)
    else:
        raise FileNotFoundError('Map: file not provided; please use --map option')

    out_filename = None
    if args.out:
        out_filename = Path(args.out)
    else:
        raise FileNotFoundError('Out filename: file not provided; please use --out option')
    
    out_filename_template = out_filename.with_stem(out_filename.stem+'')
    out_filename_include_addin = out_filename.with_stem(out_filename.stem+'.INCLUDE')
    out_filename = None
    config['out_filename_template_fullpath'] = '{s}'.format(s=Path(out_filename_template).resolve())
    config['out_filename_include_addin_fullpath'] = '{s}'.format(s=Path(out_filename_include_addin).resolve())
    config['out_filename_template_filenamepart'] = '{s}'.format(s=Path(out_filename_template).name)
    config['out_filename_include_addin_filenamepart'] = '{s}'.format(s=Path(out_filename_include_addin).name)

    # if config['need_write_template']:
        
    #     print('{script_name}: working, generating mrs script, template'.format(script_name=script_name))
    #     result_template = write_mrs.generate_savprep_mrs_template(config)

    #     print('{script_name}: saving as "{fname}"'.format(fname=out_filename_template,script_name=script_name))
    #     with open(out_filename_template, "wb") as outfile:
    #         outfile.write(codecs.BOM_UTF8)
    #     with open(out_filename_template, "a",encoding='utf-8') as outfile:
    #         outfile.write(result_template)

    print('{script_name}: working, generating mrs script, include addin'.format(script_name=script_name))
    result_include_addin = write_mrs.generate_savprep_mrs_include_part(mdd,map_df,config)

    print('{script_name}: saving as "{fname}"'.format(fname=out_filename_include_addin,script_name=script_name))
    with open(out_filename_include_addin, "wb") as outfile:
        outfile.write(codecs.BOM_UTF8)
    with open(out_filename_include_addin, "a",encoding='utf-8') as outfile:
        outfile.write(result_include_addin)

    time_finish = datetime.now()
    print('{script_name}: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start,script_name=script_name))




def cli():
    programs_to_run = {
        'write_mrs': cli_program_produce_savprep_mrs,
        'update_map': cli_program_update_map,
    }
    try:
        parser = argparse.ArgumentParser(
            description="mdmtools AA Revival utility",
            prog='mddtools-aarevival'
        )
        parser.add_argument(
            #'-1',
            '--program',
            choices=dict.keys(programs_to_run),
            type=str,
            required=True
        )
        args, _ = parser.parse_known_args()
        if args.program:
            program = '{arg}'.format(arg=args.program)
            if program in programs_to_run:
                programs_to_run[program]()
            else:
                raise AttributeError('program to run not recognized: {program}'.format(program=args.program))
        else:
            print('program to run not specified')
            raise AttributeError('program to run not specified')

    except Exception as e:
        # the program is designed to be user-friendly
        # that's why we reformat error messages a little bit
        # stack trace is still printed (I even made it longer to 20 steps!)
        # but the error message itself is separated and printed as the last message again

        # for example, I don't write "print('File Not Found!');exit(1);", I just write "raise FileNotFoundError()"
        print('')
        print('Stack trace:')
        print('')
        traceback.print_exception(e,limit=20)
        print('')
        print('')
        print('')
        print('Error:')
        print('')
        print('{e}'.format(e=e))
        print('')
        exit(1)



if __name__ == '__main__':
    cli()
