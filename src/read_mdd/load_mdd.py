

from pathlib import Path




import_er = None
try:
    import win32com.client
except Exception as e:
    import_er = e



class ReadMDDError(Exception):
    """Read MDD error"""

def open_mdd_document(mdd_path):
    try:
        if import_er:
            raise import_er # we only fail here, we don't fail if we don't need to read MDD
        mdmDocument = win32com.client.Dispatch("MDM.Document")
        # openConstants_oNOSAVE = 3
        openConstants_oREAD = 1
        # openConstants_oREADWRITE = 2
        # print('Oopening MDM document using method "open": "{path}"'.format(path=mdd_path))
        # we'll check that the file exists so that the error message is more informative - otherwise you see a long stack of messages that do not tell much
        if not(Path(mdd_path).is_file()):
            raise FileNotFoundError('File not found: {fname}'.format(fname=mdd_path))
        mdmDocument.Open( mdd_path, "", openConstants_oREAD )
        return mdmDocument
    except Exception as e:
        raise ReadMDDError('Failed to open MDD: {e}'.format(e=e)) from e

