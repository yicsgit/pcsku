import pandas as pd
import tkinter as tk
from tkinter import filedialog
import pymysql
def update_change_table(change_file_path):
    try:
        insert_df = pd.read_excel(change_file_path,sheet_name="改款对应表",usecols=["售卖的半成品编码","改款品牌","替换的新编码"])
        # 直接检查两列组合的重复项，不新增列
        duplicates = insert_df.duplicated(subset=['售卖的半成品编码', '改款品牌'], keep=False)
        if duplicates.any():
            # 提取重复项
            print(insert_df[duplicates])
            raise SystemExit("有重复项款，程序已结束")
        insert_df = [tuple(row) for _, row in insert_df.iterrows()]
        # -----------------------------------------------
        conn = pymysql.connect(host='localhost', user='root', password='123456', database='tz_sql')  # charset='utf8'
        cursor = conn.cursor()
    except Exception as e:
        print(f"错误:{e}")
    try:
        table_name = "补充表_改款"
        cursor.execute("TRUNCATE TABLE " + table_name)  # 先清空数据表
        insert_sql = 'INSERT INTO ' + table_name + ' VALUES (%s,%s,%s)'
        cursor.executemany(insert_sql, insert_df)  # 再新增数据内容
        print("改款表格,更新完成")
        conn.commit()
    except Exception as e:
        print(f"错误:{e}")
        conn.rollback()
    cursor.close()
    conn.close()


# 创建Tkinter根窗口（隐藏）
root = tk.Tk()
root.withdraw()
# 选择Excel文件
file_path = filedialog.askopenfilename(
    title="选择Excel文件",
    filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
)
if not file_path:
    print("未选择文件")
else:
    update_change_table(file_path)