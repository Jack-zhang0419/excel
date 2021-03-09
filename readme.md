# Excel Util

## prerequisites

install [python 3.x](https://www.python.org/) on your operating system.

```sh
pip install -r requirements.txt
```

## merge

make sure to create folder "to_merge" in the root folder, and put directories under "to_merge" folder

such as "to_merge\1" or "to_merge\2"

## sum

make sure to create folder "to_sum" in this folder, and put "sample.xlsx" under "to_sum" folder

such as "to_sum\sample.xls or sample.xlsx"

## combine

create folder "to_combine" in the root folder

copy "A0-A9.xlsx" "B0-B9.xlsx" to "to_combine" folder

## how to run

```sh
python index.py
```

## future features

- excel file name supports more than 9, such as A010 or A10
- support more than two sheets, such as C1
- support not merged column type in first column
- performance enhancements:
  - only load column range once

## reference link

- [python合并多个EXCEL表](https://www.jianshu.com/p/664b52d6933e)
- [合并excel重复行](https://www.jianshu.com/p/26f93146d564)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/index.html)
