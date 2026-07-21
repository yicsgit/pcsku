import pandas as pd
import tkinter as tk
from tkinter import filedialog
def select_file(): #
    # 创建隐藏的Tk根窗口
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 打开目录选择对话框
    directory_path = filedialog.askopenfilename(
        title='请选择目标目录',
        initialdir='/',  # 初始目录，可根据需要修改
    )

    # 处理用户取消选择的情况
    if not directory_path:
        print("未选择目录")
        return None

    print(f"已选择目录: {directory_path}")
    return directory_path
path = select_file()
result_df = pd.read_excel(path)
split_row = 30000
if len(result_df) >split_row :
    n = round(len(result_df)/split_row)
    end_idx = -1
    for i in range(n+1):
        start_idx = end_idx + 1
        end_idx = (i + 1) * split_row
        if (i + 1) * split_row>=len(result_df):
            end_idx = len(result_df)
        else:
            end_sku = result_df.iloc[end_idx,0]
            end_next_sku = result_df.iloc[end_idx + 1,0]
            while result_df.iloc[end_idx,0]==result_df.iloc[end_idx+1,0]:
                end_idx = end_idx + 1
        sub_df = result_df.iloc[start_idx:end_idx+1]
        sub_df.to_excel(rf'{path}-分割{i+1}导入表.xlsx', index=False)
print('切割完成')