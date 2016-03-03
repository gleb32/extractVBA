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
import os.path
import win32com.universal as universal
from win32com.client import Dispatch
from enum import IntEnum
from contextlib import contextmanager


ROOT_PATH = "C:\\gitlab\\temp"
file1 = "MOCVD DBTPHJ v1.accdb"
file2 = "MOCVD Equipment Report.xlsm"


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


@contextmanager
def close_workbook(workbook):
    """ Closes the workbook when finished, even on error """
    try:
        yield workbook
    finally:
        # https://msdn.microsoft.com/en-us/library/office/ff838613.aspx
        workbook.Close(False)   # Close the workbook without saving changes.


@contextmanager
def close_access_db(access_app):
    """
    Closes the access database when finished, even on error.

    https://msdn.microsoft.com/en-us/library/office/ff836850.aspx
    """
    try:
        yield access_app
    finally:
        access_app.CloseCurrentDatabase()


@contextmanager
def quit_excel(excel_com_obj):
    try:
        yield excel_com_obj
    finally:
        excel_com_obj.Quit


@contextmanager
def quit_access(access_app):
    try:
        yield access_app
    finally:
        access_app.Quit

def save_component(text, vbname, ext):
    """
    Save the extracted VBA code to a file.
    """
    write_path = os.path.join(ROOT_PATH, vbname + ext)
    print("Writing to `{}`".format(write_path))
    with open(write_path, 'w', newline='\n') as openf:
        openf.write(text)


def extract_component(component):
    """
    Exract the component information from the component COM object.
    """
    vb_name = component.Name
    vb_type = component.Type
    vb_code_module = component.CodeModule
    try:
        vb_src = vb_code_module.Lines(1, vb_code_module.CountOfLines)
    except universal.com_error:
        # it's likely that this component doesn't have any code
        # TODO: Better error handling here.
        # pywintypes.com_error: (-2147352567, 'Exception occurred.',
        #     (0, '', 'Invalid procedure call or argument',
        #      'C:\\PROGRA~1\\COMMON~1\\MICROS~1\\VBA\\VBA7\\1033\\VbLR6.chm',
        #      1000005, -2147024809), None)
        vb_src = None
    else:
        pass

    return (vb_name, vb_type, vb_code_module, vb_src)


def extract_components(workbook):
    """
    extracts and saves all VBA in a given workbook.
    """
    i = 1
    while True:
        try:
            component = workbook.VBProject.VBComponents(i)
        except universal.com_error:
            # TODO: Better error handling here.
            # pywintypes.com_error: (-2147352567, 'Exception occurred.',
            #   0, 'VBAProject', 'Subscript out of range',
            #   'C:\\PROGRA~1\\COMMON~1\\MICROS~1\\VBA\\VBA7\\1033\\VbLR6.chm',
            #   1000009, -2146828279), None)
            break

        vb_name, vb_type, _, vb_src = extract_component(component)

        if vb_src:
            ext = EXTENSIONS[vb_type]
            save_component(vb_src, vb_name, ext)

        i += 1


def open_excel_workbook(fn):
    """
    Open Excel and the workbook and return the woorbook COM object
    """
    xl = Dispatch("Excel.Application")
    xl.Visible = 1
    return xl.Workbooks.Open(fn)


def open_access_db(fn):
    """
    Open Access and the database and return the Access application COM object.

    https://msdn.microsoft.com/en-us/library/office/ff837226.aspx
    """
    access_app = Dispatch("Access.Application")
#    access_app.Visible = 1
    try:
        access_app.OpenCurrentDatabase(fn)
    except universal.com_error as err:
        # possibly already have the database open
        # TODO: better error handling here.
        # pywintypes.com_error: (-2147352567, 'Exception occurred.',
        #     (0, None, 'You already have the database open.', None,
        #      -1, -2146820421), None)
        print(err)
    return access_app


def main():
    pth = os.path.join(ROOT_PATH, file2)

    with close_workbook(open_excel_workbook(pth)) as openwb:
        extract_components(openwb)


if __name__ == "__main__":
    main()
#    pth = os.path.join(ROOT_PATH, file1)
#    with close_access_db(open_access_db(pth)) as opendb:
#        pass
