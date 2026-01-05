
from pathlib import Path
import shutil


class BackupManager:
    def __init__(self,program_config={},config={}):
        self.program_config = program_config
        self.config = config
        pass

    def make_output_name(self,fname_to_copy):
        script_run_datetime = self.program_config['datetime']
        script_run_datetime = script_run_datetime.strftime('_%Y%m%d_%I%M%S_Z')
        return fname_to_copy.with_stem(fname_to_copy.stem+'_backup_'+script_run_datetime)

    def make_copy(self,fname_to_copy):
        script_name = self.program_config['script_name']
        print('{script_name}: creating backup of the map: {map_fname}'.format(script_name=script_name,map_fname=fname_to_copy))
        out_bkp_filename = self.make_output_name(fname_to_copy)
        from_file = Path(fname_to_copy)
        to_file = Path(out_bkp_filename)
        shutil.copy(from_file,to_file)

