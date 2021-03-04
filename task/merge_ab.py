from openpyxl.worksheet.cell_range import CellRange
from task.itask import ITask
from util.file_util import sub_dirs, filter_excel_files
from openpyxl import load_workbook, workbook
from openpyxl.worksheet.worksheet import Worksheet
import os

# https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/worksheet/cell_range.html?highlight=CellRange#


class MergeAB(ITask):
    """
    merge multiple excels into one according with file name
    """
    output_file = "merged.xlsx"

    def run(self):
        print(
            '================================================================')
        print(f"start to merge '{self.dir}' folder")

        root_folder = f"{os.getcwd()}/{self.dir}"

        self.merge(root_folder)

    def merge(self, folder):
        """
        docstring
        """
        print(
            '----------------------------------------------------------------')
        print(f'start to merge folder: {folder} ......')
        # 文件路径
        # 构建新的表格名称
        merged_file = f"{folder}/{self.output_file}"
        self.target_workbook = workbook.Workbook()
        self.target_workbook.active.title = "学校层面"
        self.target_workbook.create_sheet(title="专业群层面")
        self.target_cur_row = [1, 1]

        filter_list = filter_excel_files(folder)
        if self.output_file in filter_list:
            filter_list.remove(self.output_file)
        sorted_list = sorted(filter_list)

        for file_name in sorted_list:
            source_sheet_no, source_block_no, target_sheet_no, target_block_no = self.parse_excel_name(
                file_name)
            (copied_range, column_range,
             copied_row_count) = self.copy_excel_block(folder, file_name,
                                                       source_sheet_no,
                                                       source_block_no)
            self.paste_range_value(
                1, self.target_cur_row[target_sheet_no], column_range,
                copied_row_count + self.target_cur_row[target_sheet_no] - 1,
                self.target_workbook.worksheets[target_sheet_no], copied_range)
            self.target_workbook.active.merge_cells(CellRange("A1:A3").coord)
            self.target_workbook.active.merge_cells(CellRange("A4:A32").coord)

        self.target_workbook.save(merged_file)

    def parse_excel_name(self, file_name):
        current_file_name = file_name.split(".")[0]
        if current_file_name[0] == "A":
            source_sheet_no = target_sheet_no = 0
        else:
            source_sheet_no = target_sheet_no = 1
        source_block_no = target_block_no = int(current_file_name[1])

        if "-" in current_file_name:
            source_block_no = int(current_file_name[-1])
            if current_file_name[-2] == "A":
                source_sheet_no = 0
            else:
                source_sheet_no = 1

        return (source_sheet_no, source_block_no, target_sheet_no,
                target_block_no)

    def get_row_range(self, sheet: Worksheet, block_no: int):
        bounds = sorted([
            merged_range.bounds for merged_range in sheet.merged_cells.ranges
            if merged_range.bounds[0] == 1
        ])
        first_row_block = bounds[block_no]
        if block_no == 1:
            found = (1, first_row_block[3])
        else:
            found = (first_row_block[1], first_row_block[3])
        print(f"block {block_no} row range: {found[0]} to {found[1]}")
        return found

    def get_column_range(self, sheet: Worksheet):
        column = 1
        while type(sheet.cell(
                row=1, column=column)).__name__ == 'MergedCell' or sheet.cell(
                    row=1, column=column).value is not None:
            column = column + 1

        print(f'current sheet {sheet.title} max column: {column - 1}')
        return column - 1

    def copy_excel_block(self, folder, file_name, source_sheet_no,
                         source_block_no):
        excel = os.path.join(folder, file_name)
        source_workbook = load_workbook(filename=excel)

        column_range = self.get_column_range(
            source_workbook.worksheets[source_sheet_no])
        row_range = self.get_row_range(
            source_workbook.worksheets[source_sheet_no], source_block_no)

        return (self.copy_range_value(
            row_range[0], 1, column_range, row_range[1],
            source_workbook.worksheets[source_sheet_no]), column_range,
                row_range[1] - row_range[0] + 1)

    # Copy range of cells as a nested list
    # Takes: start cell, end cell, and sheet you want to copy from.
    # https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/worksheet/copier.html
    def copy_range_value(self, startCol, startRow, endCol, endRow, sheet):
        rangeSelected = []
        # Loops through selected Rows
        for i in range(startRow, endRow + 1, 1):
            # Appends the row to a RowSelected list
            rowSelected = []
            for j in range(startCol, endCol + 1, 1):
                rowSelected.append(sheet.cell(row=i, column=j).value)
            # Adds the RowSelected List and nests inside the rangeSelected
            rangeSelected.append(rowSelected)

        return rangeSelected

    # Paste range
    # Paste data from copyRange into template sheet
    def paste_range_value(self, startCol, startRow, endCol, endRow,
                          sheetReceiving, copiedData):
        countRow = 0

        for i in range(startRow, endRow + 1, 1):
            countCol = 0
            for j in range(startCol, endCol + 1, 1):

                sheetReceiving.cell(
                    row=i, column=j).value = copiedData[countRow][countCol]
                countCol += 1
            countRow += 1
