import pandas as pd
import tkinter as tk
from tkinter import filedialog
import pymysql
def update_component_table(component_file_path):

    insert_df = pd.read_excel(component_file_path,sheet_name="跨境英文成分",usecols=[1,4])
    # 检查重复项，不新增列
    duplicates = insert_df.duplicated(subset=["货号"], keep=False)
    if duplicates.any():
        # 提取重复项
        print(insert_df[duplicates])
        raise SystemExit("有重复项款，程序已结束")
    
    insert_df = insert_df.dropna(subset=["货号"])
    # 使用正则表达式一次性替换所有符号
    insert_df.iloc[:,1] = insert_df.iloc[:,1].str.replace('[，；。]', lambda x: {'，': ',', '；': ';', '。': '.'}[x.group()], regex=True)
    insert_df = [tuple(row) for _, row in insert_df.iterrows()]
        # -----------------------------------------------
    conn = pymysql.connect(host='localhost', user='root', password='123456', database='tz_sql')  # charset='utf8'
    cursor = conn.cursor()
    try:
        table_name = "补充表_跨境_成分"
        cursor.execute("TRUNCATE TABLE " + table_name)  # 先清空数据表
        insert_sql = 'INSERT INTO ' + table_name + ' VALUES (%s,%s)'
        cursor.executemany(insert_sql, insert_df)  # 再新增数据内容
        print("跨境英文成分表格,更新完成")
        conn.commit()
    except Exception as e:
        print(f"错误:{e}")
        conn.rollback()
    cursor.close()
    conn.close()





# # 创建Tkinter根窗口（隐藏）
# root = tk.Tk()
# root.withdraw()
# # 选择Excel文件
# file_path = filedialog.askopenfilename(
#     title="选择Excel文件",
#     filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
# )
# if not file_path:
#     print("未选择文件")
# else:
#     update_change_table(file_path)


def update_tag_table(tag_file_path):
    # 读取 Excel，pandas 会将合并单元格中除第一个外自动填充为 NaN
    df = pd.read_excel(tag_file_path, sheet_name='1',usecols=["款号","尺码","URL"])
    # 对 A 列进行向前填充，将 NaN 替换为上一个有效值
    df['款号'] = df['款号'].ffill()
    # # 检查重复项，不新增列
    # duplicates = df.duplicated(subset=["款号","尺码","URL"], keep=False)
    # if duplicates.any():
    #     # 提取重复项
    #     print(df[duplicates])
    #     raise SystemExit("有重复项款，程序已结束")
    # 删除重复值
    df = df.drop_duplicates(subset=["款号","尺码"], keep='first')
    df = df.dropna(subset=["款号","尺码","URL"])

    insert_df = [tuple(row) for _, row in df.iterrows()]
    # -----------------------------------------------
    conn = pymysql.connect(host='localhost', user='root', password='123456', database='tz_sql')  # charset='utf8'
    cursor = conn.cursor()
    try:
        table_name = "补充表_跨境_url"
        cursor.execute("TRUNCATE TABLE " + table_name)  # 先清空数据表
        insert_sql = 'INSERT INTO ' + table_name + ' VALUES (%s,%s,%s)'
        cursor.executemany(insert_sql, insert_df)  # 再新增数据内容
        print("跨境url表格,更新完成")   
        conn.commit()
    except Exception as e:
        print(f"错误:{e}")
        conn.rollback()
    cursor.close()
    conn.close()



update_component_table(r'C:\Users\Administrator\Desktop\跨境编码信息表.xlsx')
update_tag_table(r'C:\Users\Administrator\Desktop\聚水潭贴贴.xlsx')
