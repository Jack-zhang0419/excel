#!/usr/bin/python3
# -*- coding: utf-8 -*-

import pandas as pd
import os

frame = pd.read_excel(f"{os.getcwd()}/to_sum/sample.xlsx")
df = frame.groupby(['name']).sum()
df.to_excel(f"{os.getcwd()}/to_sum/sum.xlsx")
