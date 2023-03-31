# !/usr/bin/env python
# -*- coding:utf-8 -*-
# @FileName  :data_processor.py
# @Time      :2023/3/15 15:40
# @Author    :Boris_zhang
import pandas as pd
from openpyxl import load_workbook
import os


class DataProcessorExcel:
    def __init__(self, file_path):
        self.file_path = file_path
        self.sheet_name = None
        self.df = None
        self.header_row = None
        self.cols_to_drop = []
        self.sheet_name_list = None

        if not os.path.exists('cache'):
            os.makedirs('cache')
        # cache_path,是file_path路径中截取文件名称，放在cache文件夹路径下
        self.cache_path = 'cache/' + self.file_path.split('/')[-1]
        

    def process(self, sheet_name=None):
        # 读取Excel文件
        df = pd.read_excel(self.file_path, sheet_name=sheet_name)
        # 取消了合并单元格
        df.to_excel(self.cache_path, sheet_name=sheet_name, index=False)

        # 打开工作簿，获取“Sheet”工作表对象
        # 加载 Excel 文件
        workbook = load_workbook(self.cache_path)
        # 获取要操作的工作表
        worksheet = workbook[sheet_name]
        # 在第一行插入一行空行
        worksheet.insert_rows(1)
        # 将新的 Excel 文件保存
        workbook.save(self.cache_path)
        
        # 读取Excel文件
        self.df = pd.read_excel(self.cache_path, sheet_name)

        # 表中某一列全部是空值，删掉该列。
        for col in self.df.columns:
            if self.df[col].isnull().all():
                self.df = self.df.drop(col, axis=1)

        # 找出表中存在空行的列（处理上下合并的单元格）
        for col in self.df.columns:
            if self.df[col].isnull().any():
                # 将空值补充为上一个不为空的值
                self.df[col].fillna(method='ffill', inplace=True)

        # 自动检索表头所在行
        min_unnamed_count = float('inf')
        for i in range(min(len(self.df), 10)):
            row = self.df.iloc[i]
            if all(type(col) == str or pd.isna(col) for col in row):
                unnamed_count = len([col for col in row if 'Unnamed:' in str(col) or pd.isna(col)])
                if unnamed_count < min_unnamed_count:
                    self.header_row = i + 1
                    min_unnamed_count = unnamed_count

        # 更新表头为检索的表头所在行，并将之前的数据删去
        if self.header_row is not None:
            self.df = self.df.iloc[self.header_row-1:]
            self.df.columns = self.df.iloc[0]
            self.df = self.df[1:]

        # 删除列名为空或者包含'Unnamed'的列
        for col in self.df.columns:
            if col == '' or 'Unnamed' in col:
                self.df = self.df.drop(col, axis=1)

        # 自动检索出从0或者1开始自增的序号列去掉
        for col in self.df.columns:
            if self.df[col].dtype == 'int64':
                if self.df[col].is_monotonic_increasing:
                    self.cols_to_drop.append(col)
        self.df = self.df.drop(self.cols_to_drop, axis=1)

        # 删去有连续空值的行
        self.df = self.df.dropna(how='all', thresh=len(self.df.columns)-1)

        # 自动检索和df.columns一样的行，并删除
        for i, row in self.df.iterrows():
            if row[0] == self.df.columns[0]:
                if all(row == self.df.columns):
                    # print(row)
                    self.df = self.df.drop(i, axis=0)

        # df 去掉重复行
        self.df = self.df.drop_duplicates()
        
        return self.df

        # print(self.df)
    def save_processed_data(self, file_path):
        if self.df is None:
            raise Exception("需要先执行数据处理; you need processor.process() before save to file !")
        self.df.to_excel(file_path, index=False)

    # 获取文件表格的sheet_name列表
    def get_sheet_name_list(self):
        excel_file = pd.ExcelFile(self.file_path)
        self.sheet_name_list = excel_file.sheet_names
        return self.sheet_name_list
        

if __name__ == '__main__':
    # file_path = "data_demo/早读规划.xlsx"
    file_path = "data_demo/早读规划.xlsx"
    #
    processor = DataProcessorExcel(file_path)
    # res_df = processor.process(sheet_name='2022.9')
    # sheet_list = processor.get_sheet_name_list()
    # print(res_df)
    # print(res_df.columns)
    pass