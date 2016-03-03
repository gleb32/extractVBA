# -*- coding: utf-8 -*-
"""
Extracts all the VBA from an MS OFfice file into separate files for SCC.

Created on Wed Mar  2 14:33:00 2016

@author: dthor

Usage:
    extract.py

Options:
    -h --help           # Show this screen.
    --version           # Show version.

"""
# ---------------------------------------------------------------------------
### Imports
# ---------------------------------------------------------------------------
# Standard Library
import argparse
from contextlib import contextmanager
from enum import IntEnum
import os
import os.path
from win32com.client import Dispatch
from win32com.universal import com_error

# Third-party
#from docopt import docopt

# ---------------------------------------------------------------------------
### Constants
# ---------------------------------------------------------------------------
EXT_EXCEL = '.xlsm'
EXT_ACCESS = '.accdb'
EXT_WORD = '.docm'
EXT_PPT = '.pptm'
VALID_EXT = (EXT_EXCEL, EXT_ACCESS, EXT_WORD, EXT_PPT)

class VBCompType(IntEnum):
    """ From http://www.pretentiousname.com/excel_extractvba/ """
    STD_MODULE = 1
    CLASS_MODULE = 2
    MS_FORM = 3
    DOCUMENT = 100


EXTENSIONS = {VBCompType.STD_MODULE: '.bas',
              VBCompType.CLASS_MODULE: '.cls',
              VBCompType.MS_FORM: '.frm',
              VBCompType.DOCUMENT: '.txt',
              }

# ---------------------------------------------------------------------------
### Functions
# ---------------------------------------------------------------------------
def handle_com_err_code(err, allowed_codes):
    """
    Passes specific win32com error codes, raising others.

    See https://msdn.microsoft.com/en-us/library/aa264975(v=vs.60).aspx
    for error codes.

    Attributes of 'com_error':
    --------------------------
    ``err.args`` : tuple
        from BaseException. Contains all arguments
    ``err.hresult`` : int
        ``err.args[0]``. Possibly a _win32com_ error code? Seems to always
        be -2147352567
    ``err.strerror`` : str
        ``err.args[1]``. The error string (from  _win32com_?)

    ``err.excepinfo`` : tuple
        ``err.args[2]``. Seems to be everything from the COM error

    ``err.argerror`` : unknown
        Possibly ``err.args[3]``. Seems to always be `None`


    Indicies of 'err.excepinfo':
    ----------------------------
    + [0] : ??
    + [1] : ?? Sometimes empty, sometimes "MOCVD BD" (Access)
    + [2] : error text
    + [3] : some path to .../VBA/VBA7/...
    + [4] : VBA Error code + 1,000,000
    + [5] : ?? Some non-static code.
    """
    if not isinstance(allowed_codes, (list, tuple)):
        allowed_codes = (allowed_codes, )

    # win32com seems to add 1,000,000 to the VBA error code. /shrug
    vb_err_code = err.excepinfo[4] - 1000000

    if vb_err_code not in allowed_codes:
        err_str = "Error: '{}'  (VBA Error code {})"
        print(err_str.format(err.excepinfo[2], vb_err_code))
        raise err


@contextmanager
def open_workbook(workbook_file):
    """ Open the workbook and then closes the workbook when finished. """
    try:
        excel_app = Dispatch("Excel.Application")
        excel_app.Visible = 0
        wb_com_obj = excel_app.Workbooks.Open(workbook_file)
        yield wb_com_obj
    except com_error:
        # probably file not found.
        raise
    finally:
        try:
            # https://msdn.microsoft.com/en-us/library/office/ff838613.aspx
            wb_com_obj.Close(False)   # Close the workbook without saving
        except UnboundLocalError:
            pass


@contextmanager
def open_access_db(access_file):
    """
    Open Access and the database, returning the Access application COM
    object. Upon competion or error, close the open database.

    + https://msdn.microsoft.com/en-us/library/office/ff837226.aspx
    + https://msdn.microsoft.com/en-us/library/office/ff836850.aspx
    """
    try:
        access_app = Dispatch("Access.Application")
#        access_app.Visible = 1
        access_app.OpenCurrentDatabase(access_file)
        yield access_app
    except com_error as err:
        if err.excepinfo[2] == "You already have the database open.":
            print("Warning: database already open")
        else:
            raise err
    finally:
        try:
            access_app.CloseCurrentDatabase()
        except com_error as err:
            if "refers to an object that is closed" in err.excepinfo[2]:
                pass
            else:
                raise


def save_component(save_path, text, vbname, ext):
    """
    Save the extracted VBA code to a file.
    """
    write_path = os.path.join(save_path, vbname + ext)
    print("  Saving src to `{}`".format(write_path))
    with open(write_path, 'w', newline='\n') as openf:
        openf.write(text)


def extract_component(vb_component):
    """
    Exract the component information from the component COM object.
    """
    vb_name = vb_component.Name
    vb_type = vb_component.Type
    vb_code_module = vb_component.CodeModule
    try:
        vb_src = vb_code_module.Lines(1, vb_code_module.CountOfLines)
    except com_error as err:
        # it's likely that this component doesn't have any code
        vb_src = None
        handle_com_err_code(err, 5)      # Invalid procedure call
    else:
        pass

    return (vb_name, vb_type, vb_code_module, vb_src)


def extract_components(com_obj, save_path):
    """
    extracts and saves all VBA in a given COM object.
    """
    i = 1
    while True:
        try:
            project = com_obj.VBProject
        except AttributeError:
            # MS Access uses a sightly different structure
            # http://stackoverflow.com/a/27385063/1354930
            project = com_obj.VBE.VBProjects(1)

        try:
            component = project.VBComponents(i)
        except com_error as err:
            handle_com_err_code(err, 9)      # Subscript out of range
            break

        vb_name, vb_type, _, vb_src = extract_component(component)

        if vb_src:
            ext = EXTENSIONS[vb_type]
            save_component(save_path, vb_src, vb_name, ext)

        # We can't loop in a pythonic way because we can't just get a list
        # of VBComponent items.
        i += 1


def main(path=None,
         excel_only=False,
         ):
    """
    """
    if path is None:
        raise ValueError("The 'path' argument is required")

    for dirpath, dirnames, filenames in os.walk(path):
        # skip over the .git directory, removing it so we don't traverse it.
        if '.git' in dirnames:
            dirnames.remove('.git')

        # ignore all non-Office doc files and temp files ("~filename.ext")
        filenames = [f for f in filenames
                     if (f.endswith(VALID_EXT) and f[0] != '~')]

        # Loop through all of the remaining files.
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
            name, ext = os.path.splitext(filename)
            name = "_src~" + name
            save_path = os.path.join(dirpath, name)

            # create a folder for the file's code
            try:
                os.mkdir(os.path.join(dirpath, name))
            except FileExistsError:
                pass

            # Extract from the various file types.
            if ext == EXT_EXCEL:
                print("Extracting from Excel: `{}`".format(filename))
                try:
                    with open_workbook(filepath) as openwb:
                        extract_components(openwb, save_path)
                except com_error as err:
                    print("!! Error extacting from {}".format(filename))
                    continue

            elif ext == EXT_ACCESS:
                print("Extracting from Access: `{}`".format(filename))
                try:
                    with open_access_db(filepath) as opendb:
                        extract_components(opendb, save_path)
                except com_error as err:
                    print("!! Error extacting from {}".format(filename))
                    continue

            elif ext == EXT_WORD:
                print("Extracting from Word: `{}`".format(filename))
                print("FORMAT NOT YET SUPPORTED")

            elif ext == EXT_PPT:
                print("Extracting from PowerPoint: `{}`".format(filename))
                print("FORMAT NOT YET SUPPORTED")

            else:
                raise ValueError("How did you even GET here??")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("path",
                        help="the path to look for Office docs in",
                        type=str,
                        )
    parser.add_argument("-x", "--excel-only",
                        help="only work on excel *.xlsm files",
                        action="store_true",
                        )

    args = parser.parse_args()

    main(**vars(args))

