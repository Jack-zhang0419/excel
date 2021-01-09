import os


def sub_dirs(dir):
    """
    return sub folders of dir
    """
    return [x[0] for x in os.walk(dir) if x[0] != dir]


def filter_excel_files(dir):
    """
    return *.xls or *.xlsx files in dir
    """
    excel_files = []

    file_list = [x for x in os.listdir(dir)]

    for file in file_list:
        if file.endswith('.xls') or file.endswith('.xlsx'):
            excel_files.append(file)

    return excel_files
