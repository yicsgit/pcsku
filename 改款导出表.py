import pandas as pd
from tkinter import filedialog
from pathlib import Path


excel_path = filedialog.askopenfilename(
    title="选择聚水潭的源数据Excel文件",
    filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
if not excel_path:
    print("未选择文件")
    quit()
try:
    source_df = pd.read_excel(excel_path, sheet_name=0)
    print(f"成功读取文件，原始数据共 {len(source_df)} 行")
    # 1. 剔除"实体编码"列上前三位字符是100的数据
    # 确保实体编码为字符串类型再进行截取
    source_df = source_df[~source_df["组合商品实体编码"].astype(str).str[:3].eq("100")]
    # 2. 剔除"备注"列上是"供销"的数据
    source_df = source_df[~source_df["备注"].eq("供销")]
    # 3. 选择"品牌"列上包含"真维斯"或"回力"的数据
    # 使用str.contains配合正则表达式实现多关键词匹配
    brands = ["公版", "LA CHAPELLE KIDS", "史努比", "棉致", "巴布豆", "真维斯", "真维斯Life", "回力", "班尼路", "唐狮",
              "马卡乐", "小猪班纳", "小黄人", "巴帝巴帝", "匹克", "贵人鸟","锐步", "Dickies", "红豆","迪士尼","SPAO"]
    pattern = '|'.join(brands)
    source_df = source_df[source_df["品牌"].astype(str).str.contains(pattern, case=False, regex=True)]

    需要替换的品牌 = ["公版", "LA CHAPELLE KIDS", "史努比", "棉致", "巴布豆", "真维斯", "真维斯life", "回力", "班尼路", "唐狮",
              "马卡乐", "小猪班纳", "小黄人", "巴帝巴帝", "锐步", "红豆","迪士尼"]
    pattern = '|'.join(需要替换的品牌)
    source_df['品牌'] = source_df['品牌'].replace(f'(?i){pattern}', '常规品牌',regex=True)

    print(f"筛选后的数据共 {len(source_df)} 行\n")
    source_df = source_df[["组合商品编码","商品编码","品牌","数量"]]

except Exception as e:
    print(f"读取Excel文件出错: {str(e)}")
    quit()


excel_path = filedialog.askopenfilename(
    title="选择改款表Excel文件",
    filetypes=[("Excel files", "*.xlsx *.xls")])
if not excel_path:
    print("未选择文件")
    quit()
try:
    change_file_name = Path(excel_path).stem
    change_df = pd.read_excel(excel_path, sheet_name="改款对应表", usecols=["售卖的半成品编码","改款品牌","替换的新编码"])
    change_df['改款品牌'] = change_df['改款品牌'].replace('dickies','Dickies')
    duplicates = change_df.duplicated(subset=['售卖的半成品编码', '改款品牌'], keep=False)
    if duplicates.any():
        # 提取重复项
        print(change_df[duplicates])
        raise SystemExit("有重复项款，程序已结束")
    change_df.columns = ['商品编码','品牌','替换的新编码']
except Exception as e:
    print(f"读取Excel文件出错: {str(e)}")
    quit()


merger_df = pd.merge(source_df,change_df,how='left',on=['商品编码','品牌'])
merger_df['替换的新编码'] = merger_df['替换的新编码'].fillna(merger_df['商品编码'])
result_df = merger_df[["组合商品编码","替换的新编码","数量"]].rename(columns={'替换的新编码': '商品编码'})

split_row = 30000
if len(result_df) <=split_row :
    result_df.to_excel(rf'\\PCWCtz\童装\童装数据资料\temp_excel\组合\{change_file_name}-改款导入表.xlsx', index=False)
else:
    n = round(len(result_df)/split_row)
    end_idx = -1
    for i in range(n+1):
        start_idx = end_idx + 1
        end_idx = (i + 1) * split_row
        if (i + 1) * split_row>=len(result_df):
            end_idx = len(result_df)
        else:
            while result_df.loc[end_idx,"组合商品编码"]==result_df.loc[end_idx+1,"组合商品编码"]:
                end_idx = end_idx + 1
        sub_df = result_df.iloc[start_idx:end_idx+1]
        sub_df.to_excel(rf'\\PCWCtz\童装\童装数据资料\temp_excel\组合\{change_file_name}-分割{i+1}导入表.xlsx', index=False)
print('改款完成')