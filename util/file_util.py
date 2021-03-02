import os
from pathlib import PurePath


def sub_dirs(dir):
    """
    return sub folders of dir
    """
    return [x[0] for x in os.walk(dir) if x[0] != dir]


def filter_excel_files(dir):
    """
    return *.xls or *.xlsx files in dir
    """

    return [
        x for x in os.listdir(dir)
        if PurePath(x).match('*.xls') or PurePath(x).match('*.xlsx')
    ]
