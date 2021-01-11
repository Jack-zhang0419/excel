import pandas as pd
import os
from task.itask import ITask


class Sum(ITask):
    """
    docstring
    """
    output_file = "sum.xlsx"
    group_column_name = "name"
    excel_file = "sample.xlsx"

    # def __init__(self, dir):
    #     super(ITask, self).__init__(dir)

    def run(self):
        print(
            '================================================================')
        print(f"start to sum '{self.dir}' folder")

        frame = pd.read_excel(f"{os.getcwd()}/{self.dir}/{self.excel_file}")
        df = frame.groupby([self.group_column_name]).sum()
        df.to_excel(f"{os.getcwd()}/{self.dir}/{self.output_file}")

        print("done!")
