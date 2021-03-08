# Excel 合并说明

## 准备工作

- 安装最新版本的[Python](http://www.python.org)
- 安装依赖包:

```cmd
@REM 这行是注释：请cd到当前目录，例如: cd merge-excel-master

pip install -r requirements.txt
```

## 执行

- 在当前目录新文件夹("**to_combine**"), 如果存在不需要重复创建
- 把需要合并的excel文件拷贝到“to_combine”文件夹下
- excel文件的命名规则:
  - <目标SHEET><目标区块>-<源SHEET><源区块>
  - <目标SHEET> or <源SHEET>: A代表第一个SHEET, B代表第二个SHEET, 依此类推，最大支持到Z,也就是26个SHEET（A, B, C ... Z）
  - <目标区块> or <源区块>: 必须是整数，从0开始，例如：0，1，2，10，20，111等，其中0比较特殊，例如：A0代表从此文件的第一个SHEET中拷贝Header，并根据它设置第一个SHEET所有列的宽度
  - 例如：A0-A0.xlsx 意思是将此excel的第一个SHEET的第0个区块(Header)拷贝到目标excel的第一个SHEET的第一个区块(Header)
  - 当<目标SHEET><目标区块>和<源SHEET><源区块>完全一致的情况下，可缩写，例如：A1.xlsx

excel文件准备好后，在windows上双击: "**run.cmd**"

如果执行成功，会在最后一行显示“done”，同时在当前目录生成"combined.xlsx"文件

如果执行失败，在窗口中会有错误提示，请将错误信息发送给维护人员
