from task.source_excel import SourceExcel
from task.target_excel import TargetExcel
from task.itask import ITask
from util.file_util import filter_excel_files
# from openpyxl.worksheet.copier import WorksheetCopy
import os


class Combine(ITask):
    """
    merge multiple excels into one according with file name
    """
    output_file = "combined.xlsx"

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
        combined_file = f"{folder}/{self.output_file}"
        self.target = TargetExcel(combined_file)

        filter_list = filter_excel_files(folder)
        if self.output_file in filter_list:
            filter_list.remove(self.output_file)
        sorted_list = sorted(filter_list)

        worksheet_level_set = [False, False]

        for file_name in sorted_list:
            print(f"--------- start {file_name} ---------")
            source_sheet_no, source_block_no, target_sheet_no, target_block_no = self.parse_excel_name(
                file_name)
            print(
                f"working on source sheet: {source_sheet_no}, source block: {source_block_no}, target sheet: {target_sheet_no}"
            )
            self.target.switch_sheet(target_sheet_no)
            self.source = SourceExcel(folder, file_name, source_sheet_no,
                                      source_block_no)

            print(f"copy/paste data and style of excel block range")
            self.target.paste_excel_block(*self.source.copy_excel_block(),
                                          source_block_no)
            print(f"copy/paste merged cell range of excel block range")
            self.target.paste_merged_cell_range(
                *self.source.get_merged_cell_range())

            # only do once
            if worksheet_level_set[target_sheet_no] is False:
                print(f"do worksheet level setting: {target_sheet_no}")
                self.target.set_worksheet_dimensions()

                worksheet_level_set[target_sheet_no] = True

            self.target.set_start_row()
            self.source = None

        self.target.save()

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
