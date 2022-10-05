# -*- coding: utf-8 -*-
"""
Created on Tue Oct 05 10:04:02 2022

@author: Administrator
"""
import shutil
import os
from openpyxl import load_workbook
# import pandas as pd


def main():
    folderpath = input("筛选文件夹:\n")
    folderpath = folderpath.replace("\'", "").replace("\"", "")
    filterfile = input("过滤的xlsx文件:\n")
    filterfile = filterfile.replace("\'", "").replace("\"", "")
    # df = pd.read_excel(filterfile, sheet_name=0)
    # df = df.iloc[:, 0]
    # print(df)
    wb = load_workbook(filterfile)
    sheetnames = wb.sheetnames
    ws = wb[sheetnames[0]]  # index为0为第一张表
    parentpath = folderpath + "\\"
    print(parentpath)
    filterfolder = parentpath + "过滤文件夹"
    if not os.path.isdir(filterfolder):
        os.mkdir(filterfolder)
    try:
        # for i in range(len(df)):
        for i in range(1, ws.max_row + 1):
            # if os.path.isdir(parentpath + str(df[i])):
            if os.path.isdir(parentpath + str(ws.cell(i, 1).value)):
                # a = shutil.move(parentpath + str(df[i]), filterfolder)
                a = shutil.move(parentpath + str(ws.cell(i, 1).value), filterfolder)
                print(a)
    except Exception as e:
        # print(parentpath + str(df[i]) + ":移动异常")
        print(parentpath + str(ws.cell(i, 1).value) + ":移动异常")
        print(repr(e))
    input("过滤完成, 按回车结束")


if __name__ == "__main__":
    main()
