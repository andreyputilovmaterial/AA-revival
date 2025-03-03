
import pandas as pd


from pathlib import Path


try:
    from zoneinfo import ZoneInfo
    import time
    get_localzone_name = None
    try:
        from tzlocal import get_localzone_name
    except:
        pass # if it's not installed, I will not require it, it's not essential
except:
    ZoneInfo = None
    pass # older python




from . import util_dataframe_wrapper
from .column_definitions import sheet_overview
from .column_definitions import sheet_mdddata_variables
from .column_definitions import sheet_mdddata_categories







CONFIG_TOOL_NAME = 'IPS Tools - Author Specifications Revival v0.000'




def build_df(mdd,config):

    # if df:
    #     df.drop(df.index,inplace=True)

    df = util_dataframe_wrapper.PandasDataframeWrapper(sheet_overview.columns).to_df()

    mdmdoc = mdd.mdmDocument

    mdd_filename_part = Path(config['mdd_filename']).stem
    mdd_reload_datetime = config['datetime']

    config_datetime_timezone_info_included = False
    try:
        mdd_reload_datetime = mdd_reload_datetime.strftime('%m/%d/%Y %I:%M:%S %p %Z')
        if ZoneInfo:
            try:
                mdd_reload_datetime_naive = config['datetime']
                try:
                    system_tz = ZoneInfo(time.tzname[0])
                except Exception as e:
                    if get_localzone_name:
                        system_tz = ZoneInfo(get_localzone_name())
                    else:
                        raise e
                mdd_reload_datetime_local = mdd_reload_datetime_naive.replace(tzinfo=system_tz)
                mdd_reload_datetime_local = mdd_reload_datetime_local.strftime('%m/%d/%Y %I:%M:%S %p %Z')
                mdd_reload_datetime = mdd_reload_datetime_local
                config_datetime_timezone_info_included = True
            except:
                pass
    except:
        pass
    if not config_datetime_timezone_info_included:
        mdd_reload_datetime = '{dt} (time configured as system time on a local machine where the script was running, without timezone information)'.format(dt=mdd_reload_datetime)
    
    mdd_edit_datetime = None
    try:
        savelogs = [ log for log in mdmdoc.SaveLogs ]
        mdd_edit_datetime = savelogs[-1].Date
        mdd_edit_datetime = mdd_edit_datetime.strftime('%m/%d/%Y %I:%M:%S')
        mdd_edit_datetime = '{dt} (time configured as system time on a local machine where processing script was running, without timezone information)'.format(dt=mdd_edit_datetime)
    except Exception as e:
        mdd_edit_datetime = '{er}'.format(er=e)

    df.loc['Tool'] = [CONFIG_TOOL_NAME]
    df.loc['MDD'] = [mdd_filename_part]
    df.loc['Validation'] = [
        '=IF(AND(AND({range_validation_variables}),AND({range_validation_categories})),"No issues","Failed")'.format(
            range_validation_variables = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = sheet_mdddata_variables.sheet_name,
                col = sheet_mdddata_variables.column_letters['col_validation'],
            ),
            range_validation_categories = '\'{sheet_name}\'!${col}$2:${col}$999999'.format(
                sheet_name = sheet_mdddata_categories.sheet_name,
                col = sheet_mdddata_categories.column_letters['col_validation'],
            ),
        ), 
    ]
    df.loc['MDD path'] = [config['mdd_filename']]
    df.loc['MDD metadata last read, date and time'] = [mdd_reload_datetime]
    df.loc['MDD modification date and time'] = [mdd_edit_datetime]
    df.loc['JobNumber property from MDD'] = [mdmdoc.Properties['JobNumber']]
    df.loc['StudyName propertty from MDD'] = [mdmdoc.Properties['StudyName']]
    df.loc['Client property from MDD'] = [mdmdoc.Properties['Client']]
    return df


