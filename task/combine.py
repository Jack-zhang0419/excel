from task.combine_util.source_excel import SourceExcel
from task.combine_util.target_excel import TargetExcel
from task.itask import ITask
from util.file_util import filter_excel_files
# from openpyxl.worksheet.copier import WorksheetCopy
from task.combine_util.name_parser import parse_file_names
import os


class Combine(ITask):
    """
    merge multiple excels into one according with file name
    """
    OUTPUT_FILE_NAME = "combined.xlsx"

    def run(self):
        print(
            '================================================================')

        root_folder = f"{os.getcwd()}/{self.dir}"
        print(f"start to combine '{root_folder}' folder")

        self.combine(root_folder)

    def combine(self, folder):
        """
        copy range of cells from varies excels and paste into one by sequence
        """
        combined_file = f"{folder}/{self.OUTPUT_FILE_NAME}"
        self.target = TargetExcel(combined_file)

        filter_list = filter_excel_files(folder)
        if self.OUTPUT_FILE_NAME in filter_list:
            filter_list.remove(self.OUTPUT_FILE_NAME)

        sorted_list = parse_file_names(filter_list)
        for parsed in sorted_list:
            target_sheet_no, sequence_no, source_sheet_no, source_block_no, orginal_file_name = parsed
            print(
                f"--------- start to process {orginal_file_name} and sequence_no: {sequence_no}---------"
            )
            print(
                f"working on source sheet: {source_sheet_no}, source block: {source_block_no}, target sheet: {target_sheet_no}"
            )
            self.target.switch_sheet(target_sheet_no)
            self.source = SourceExcel(folder, orginal_file_name,
                                      source_sheet_no, source_block_no)
            # only do once per sheet
            if self.target.worksheet_level_set[target_sheet_no] is False:
                print(f"do worksheet level setting once: {target_sheet_no}")
                self.target.set_worksheet_column_dimensions(
                    self.source.get_column_dimensions())
                self.target.worksheet.title = self.source.worksheet.title

                self.target.worksheet_level_set[target_sheet_no] = True

            # self.source.calculate_row_range()
            print(f"copy/paste data and style of block range")
            self.target.paste_excel_block(*self.source.copy_excel_block())
            self.target.set_row_dimensions(*self.source.get_row_dimensions())
            print(f"copy/paste merged cell range of block range")
            self.target.paste_merged_cell_range(
                *self.source.get_merged_cell_range())

            self.target.set_start_row()
            self.target.increase_block_no()

            self.source = None

        self.target.save()
        print(
            '================================================================')
        print("Done")
