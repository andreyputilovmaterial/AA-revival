from datetime import datetime, timezone
import argparse
from pathlib import Path
import traceback

import pandas as pd




if __name__ == '__main__':
    # run as a program
    import produce_savprep_mrs
    import read_mdd.list_variables as read_mdd
    import read_map
elif '.' in __name__:
    # package
    from . import produce_savprep_mrs
    from .read_mdd import list_variables as read_mdd
    from . import read_map
else:
    # included with no parent package
    import produce_savprep_mrs
    import read_mdd.list_variables as read_mdd
    import read_map





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
    args, _ = parser.parse_known_args()
    
    variables = None
    if args.mdd:
        mdd_filename = Path(args.mdd)
        mdd_filename = '{fname}'.format(fname=mdd_filename.resolve())
        if not(Path(mdd_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=mdd_filename))
        print('{script_name}: reading mdd "{fname}"'.format(fname=mdd_filename,script_name=script_name))
        variables = read_mdd.get_variables(mdd_filename)
    else:
        raise FileNotFoundError('MDD: file not provided; please use --mdd option')

    map_df = None
    if args.map:
        map_filename = Path(args.map)
        map_filename = '{fname}'.format(fname=map_filename.resolve())
        if not(Path(map_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=map_filename))
        print('{script_name}: reading map "{fname}"'.format(fname=map_filename,script_name=script_name))
        map_df = read_map.load(map_filename)
    else:
        raise FileNotFoundError('Map: file not provided; please use --map option')

    config = {}
    config['datetime'] = time_start

    out_filename = None
    if args.out:
        out_filename = Path(args.out)
    else:
        raise FileNotFoundError('Out filename: file not provided; please use --out option')

    print('{script_name}: working, generating mrs script'.format(script_name=script_name))
    result = produce_savprep_mrs.generate_savprep_mrs_scripts(variables,map_df,config)
    
    print('{script_name}: saving as "{fname}"'.format(fname=out_filename,script_name=script_name))
    with open(out_filename, "w",encoding='utf-8') as outfile:
        outfile.write(result)

    time_finish = datetime.now()
    print('{script_name}: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start,script_name=script_name))




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
    args, _ = parser.parse_known_args()
    
    mdd_filename = None
    variables = None
    if args.mdd:
        mdd_filename = Path(args.mdd)
        mdd_filename = '{fname}'.format(fname=mdd_filename.resolve())
        if not(Path(mdd_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=mdd_filename))
        print('{script_name}: reading mdd "{fname}"'.format(fname=mdd_filename,script_name=script_name))
        variables = read_mdd.get_variables(mdd_filename)

    map_filename = None
    map_df = None
    if args.map:
        map_filename = Path(args.map)
        map_filename = '{fname}'.format(fname=map_filename.resolve())
        if not(Path(map_filename).is_file()):
            raise FileNotFoundError('file not found: {fname}'.format(fname=map_filename))
        print('{script_name}: reading map "{fname}"'.format(fname=map_filename,script_name=script_name))
        map_df = read_map.load(map_filename)
    
    config = {}
    config['datetime'] = time_start

    if args.config_fill_analysisvalues:
        config['map_options'] = {
            'process_analysis_values': args.config_fill_analysisvalues,
        }
    
    config['mdd_file_provided'] = not not mdd_filename
    config['map_file_provided'] = not not map_filename
    config['mdd_filename'] = mdd_filename if mdd_filename else ''
    config['map_filename'] = map_filename if map_filename else ''

    if not mdd_filename:
        if not map_filename:
            raise FileNotFoundError('MDD not map file are provided. You have to pass MDD file name and/or map file name. If MDD file name is not provided, it can be read from the map. If map file is not provided, it will be created with a default name. But you have to set of of those, --mdd or --map, or both.')
        print('{script_name}: MDD name not provided, reading MDD path from the map'.format(script_name=script_name))
        with pd.ExcelFile(map_filename,engine='openpyxl') as xls:
            df = xls.parse(sheet_name='Overview', header=0, index_col=0, keep_default_na=False).fillna('')
            try:
                mdd_filename = df.at['MDD path','Value']
            except Exception as e:
                raise FileNotFoundError('Can\'t read MDD file name from map: {e}'.format(e=e)) from e
            mdd_filename = Path(mdd_filename)
            mdd_filename = '{fname}'.format(fname=mdd_filename.resolve())
            if not(Path(mdd_filename).is_file()):
                raise FileNotFoundError('file not found: {fname}'.format(fname=mdd_filename))
            print('{script_name}: reading MDD "{fname}"'.format(fname=mdd_filename,script_name=script_name))
            variables = read_mdd.get_variables(mdd_filename)
            config['mdd_filename'] = mdd_filename if mdd_filename else ''

    out_filename = map_filename if map_filename else ( ( Path(mdd_filename) if Path(mdd_filename).is_dir() else Path(mdd_filename).parents[0] ) if mdd_filename else Path('.') ) / 'AnalysisAuthorRevival.xlsx'
    out_filename = Path(out_filename)
    out_filename = out_filename.with_stem(out_filename.stem+'.updated') # TODO: debug, code for testing only, remove when done with testing
    assert out_filename, 'out filename is still missing, please check the code'

    print('{script_name}: working, updating the map'.format(script_name=script_name))
    result_df = read_map.update_map(variables,map_df,config)
    
    print('{script_name}: saving as "{fname}"'.format(fname=out_filename,script_name=script_name))
    read_map.write_to_file(result_df,out_filename)

    time_finish = datetime.now()
    print('{script_name}: finished at {dt} (elapsed {duration})'.format(dt=time_finish,duration=time_finish-time_start,script_name=script_name))




def cli():
    programs_to_run = {
        'produce_savprep_mrs': cli_program_produce_savprep_mrs,
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
