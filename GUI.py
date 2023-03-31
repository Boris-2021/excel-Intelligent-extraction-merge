# !/usr/bin/env python
# -*- coding:utf-8 -*-
# @FileName  :GUI.py
# @Time      :2023/3/27 11:49
# @Author    :Boris_zhang

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from data_processor import DataProcessorExcel
import pandas as pd
import os


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.file_path = None
        self.sheet_name_list = ['暂无', '暂无']
        self.sheet_name_label = None
        self.merge_df = None
        self.res_df = None
        self.create_widgets()

        # Bind the function to the window's closing event
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

    def clear_cache_folder(self):
        # Get the path of the cache folder
        cache_folder = "./cache"

        # Check if the cache folder exists
        if os.path.exists(cache_folder):
            # Iterate over the files in the cache folder
            for file in os.listdir(cache_folder):
                # Get the full path of the file
                file_path = os.path.join(cache_folder, file)

                # Check if the file is a file (not a directory) and delete it
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    print(f"Deleted file: {file_path}")
            os.rmdir(cache_folder)

    def on_closing(self):
        # Clear the cache folder
        self.clear_cache_folder()

        # Close the GUI
        self.master.destroy()

    def merge_data_toexcel(self, path):
        # Check if merge_df exists
        if self.merge_df is None:
            print("No data to save.")
            return

        # Save the merged data to the specified path as an Excel file
        try:
            with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
                self.merge_df.to_excel(writer, sheet_name="Merged Data", index=False)
            print("Data saved to", path)
        except Exception as e:
            print("Error:", e)

    def get_save_path(self):
        # Open a file dialog to select a save path
        self.save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel file", "*.xlsx")])
        if self.save_path:
            self.merge_data_toexcel(self.save_path)

    def create_widgets(self):
        # create left frame
        self.left_frame = tk.Frame(self)
        self.left_frame.pack(side="left", fill="y", padx=10, pady=10)

        # Select File button
        self.file_path_button = tk.Button(self.left_frame, text="选择文件", command=self.get_file_path)
        self.file_path_button.pack(side="top", padx=10, pady=10)

        # Sheet Name dropdown
        self.sheet_name_label = tk.Label(self.left_frame, text="Sheet Name(子表名称):")
        self.sheet_name_label.pack(side="top", padx=10, pady=5)

        self.sheet_name_var = tk.StringVar(self.left_frame)
        self.sheet_name_var.set(self.sheet_name_list[0])
        self.sheet_name_dropdown = tk.OptionMenu(self.left_frame, self.sheet_name_var, *self.sheet_name_list)
        self.sheet_name_dropdown.config(width=15)
        self.sheet_name_dropdown.pack(side="top", padx=5, pady=5)

        # Start button
        self.start_button = tk.Button(self.left_frame, text="开始抽取", command=self.start_processing)
        self.start_button.pack(side="top", padx=10, pady=5)

        # Merge button
        self.merge_button = tk.Button(self.left_frame, text="并入", command=self.merge_data)
        self.merge_button.pack(side="top", padx=10, pady=5)

        # Save button
        self.save_button = tk.Button(self.left_frame, text="保存文件", command=self.get_save_path)
        self.save_button.pack(side="bottom", padx=10, pady=5)

        # create right frame
        self.right_frame = tk.Frame(self)
        self.right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)


        # Processed Data Treeview
        self.res_df_view = ttk.Treeview(self.right_frame)
        self.res_df_view.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        # Merged Data label
        self.merge_label = tk.Label(self.right_frame, text="并入后的数据 ↥↥↥↥↥ :")
        self.merge_label.pack(side="bottom", padx=10, pady=5)

        # Merged Data Treeview
        self.merge_df_view = ttk.Treeview(self.right_frame)
        self.merge_df_view.pack(side="bottom", fill="both", expand=True, padx=10, pady=10)

    def get_file_path(self):
        self.file_path = tk.filedialog.askopenfilename()
        processor = DataProcessorExcel(self.file_path)
        self.sheet_name_list = processor.get_sheet_name_list()
        self.sheet_name_var.set(self.sheet_name_list[0])
        self.sheet_name_dropdown['menu'].delete(0, 'end')
        for sheet_name in self.sheet_name_list:
            self.sheet_name_dropdown['menu'].add_command(label=sheet_name, command=tk._setit(self.sheet_name_var, sheet_name))

    def start_processing(self):
        sheet_name = self.sheet_name_var.get()
        # print(sheet_name)
        processor = DataProcessorExcel(self.file_path)
        self.res_df = processor.process(sheet_name=sheet_name)
        # print(self.res_df)

        # Clear the previous data in the TreeView widget
        self.res_df_view.delete(*self.res_df_view.get_children())

        # Insert the column headings
        self.res_df_view["columns"] = list(self.res_df.columns)
        for col in self.res_df_view["columns"]:
            self.res_df_view.heading(col, text=col)

        # Insert the data rows
        for i, row in self.res_df.iterrows():
            self.res_df_view.insert("", "end", text=i, values=list(row))

    def merge_data(self):
        if self.merge_df is None:
            self.merge_df = self.res_df
        else:
            self.merge_df = pd.concat([self.merge_df, self.res_df], axis=0)
        self.merge_df.reset_index(drop=True, inplace=True)

        # Clear the previous data in the TreeView widget
        self.merge_df_view.delete(*self.merge_df_view.get_children())

        # Insert the column headings
        self.merge_df_view["columns"] = list(self.merge_df.columns)
        for col in self.merge_df_view["columns"]:
            self.merge_df_view.heading(col, text=col)

        # Insert the data rows
        for i, row in self.merge_df.iterrows():
            self.merge_df_view.insert("", "end", text=i, values=list(row))


root = tk.Tk()
app = Application(master=root)
app.mainloop()


if __name__ == "__main__":
    pass
