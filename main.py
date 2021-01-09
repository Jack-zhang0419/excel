import pandas as pd
import os


def merge(sub_dir):
    """
    docstring
    """
    print('================================================================')
    print(f'start to merge {sub_dir} ...')
    # 文件路径
    # 构建新的表格名称
    merged_file = sub_dir + '/merged.xlsx'
    # 找到文件路径下的所有表格名称，返回列表
    file_list = os.listdir(sub_dir)
    new_list = []

    for file in file_list:
        if file != 'merged.xlsx' and (file.endswith('.xls')
                                      or file.endswith('.xlsx')):
            # 重构文件路径
            file_path = os.path.join(sub_dir, file)
            # 将excel转换成DataFrame
            print(file_path)
            dataframe = pd.read_excel(file_path)
            # 保存到新列表中
            new_list.append(dataframe)

    # 多个DataFrame合并为一个
    df = pd.concat(new_list)
    # 写入到一个新excel表中
    df.to_excel(merged_file, index=False)
    print(f'merged to {merged_file}')


def main():
    inputFolder = f"{os.getcwd()}/input/"
    # print(inputFolder)
    subdirs = []
    for subdir in [x[0] for x in os.walk(inputFolder)]:
        if subdir != inputFolder:
            subdirs.append(subdir)

    for subdir in subdirs:
        merge(subdir)


if __name__ == "__main__":
    main()
