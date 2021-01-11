from task.itask import ITask
from util.file_util import sub_dirs, filter_excel_files
import pandas as pd
import os


class Merge(ITask):
    """
    merge multiple excels into one
    """
    output_file = "merged.xlsx"

    def run(self):
        print(
            '================================================================')
        print(f"start to merge '{self.dir}' folder")

        root_folder = f"{os.getcwd()}/{self.dir}"

        for subdir in sub_dirs(root_folder):
            self.merge(subdir)

    def merge(self, sub_dir):
        """
        docstring
        """
        print(
            '----------------------------------------------------------------')
        print(f'start to merge sub folder: {sub_dir} ......')
        # 文件路径
        # 构建新的表格名称
        merged_file = f"{sub_dir}/{self.output_file}"
        # 找到文件路径下的所有表格名称，返回列表
        new_list = []

        filter_list = filter_excel_files(sub_dir)

        for file in filter_list:
            if not file.endswith(self.output_file):
                # 重构文件路径
                file_path = os.path.join(sub_dir, file)
                # 将excel转换成DataFrame
                print(f'merging {file_path} ...')
                dataframe = pd.read_excel(file_path)
                # 保存到新列表中
                new_list.append(dataframe)

        # 多个DataFrame合并为一个
        df = pd.concat(new_list)
        # 写入到一个新excel表中
        df.to_excel(merged_file, index=False)
        print(f'merged to {merged_file}')
