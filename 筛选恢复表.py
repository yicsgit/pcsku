import pandas as pd
import tkinter as tk
from tkinter import filedialog,messagebox
import openpyxl

def select_and_filter_file(brand_mode):
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
        return

    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        print(f"成功读取文件，原始数据共 {len(df)} 行")

        # 1. 剔除"实体编码"列上前三位字符是100的数据
        # 确保实体编码为字符串类型再进行截取
        df = df[~df["组合商品实体编码"].astype(str).str[:3].eq("100")]

        # 2. 剔除"备注"列上是"供销"的数据
        df = df[~df["备注"].eq("供销")]

        # 3. 选择"品牌"列上包含"真维斯"或"回力"的数据
        # 使用str.contains配合正则表达式实现多关键词匹配
        if brand_mode == 0 :
            brands = ["公版","LA CHAPELLE KIDS","史努比","棉致","巴布豆","真维斯","真维斯Life","回力","班尼路","唐狮","马卡乐","小猪班纳","小黄人","巴帝巴帝","匹克","贵人鸟","Dickies","红豆","迪士尼"]
        elif brand_mode == 1 :
            brands = ["公版","LA CHAPELLE KIDS","史努比","棉致","巴布豆","真维斯","真维斯Life","回力","班尼路","唐狮","马卡乐","小猪班纳","小黄人","巴帝巴帝","红豆","迪士尼"]
        else:
            print('模式错误')
            quit()
        pattern = '|'.join(brands)
        df = df[df["品牌"].astype(str).str.contains(pattern,case=False, regex=True)]

        print(f"筛选后的数据共 {len(df)} 行")


        def process_2code(code):
                left_code = code[:-5]
                last_five = code[-5:]
                # 替换为特定格式：-前4位-最后1位（例如10501 → -1050-1）
                return f"{left_code}-{last_five[:4]}-{last_five[-1]}"
        # 新增列初始化
        df['供应商商品编码'] = df['组合商品实体编码'].astype(str)
        # 条件判断与转换
        mask = df['虚拟分类'] == '两件组合'
        df.loc[mask, '供应商商品编码'] = df.loc[mask, '组合商品实体编码'].apply(process_2code)
        # def process_code(row):
        #     if row["虚拟分类"] == "两件组合":
        #         # 获取编码后5位
        #         last_five = row["组合商品实体编码"][-5:]
        #         # 替换为特定格式：-前4位-最后1位（例如10501 → -1050-1）
        #         return f"-{last_five[:4]}-{last_five[-1]}"
        #     else:
        #         # 非两件组合时直接返回原编码
        #         return row["组合商品实体编码"]
        # # 新增列并应用处理函数
        # df["供应商商品编码"] = df.apply(process_code, axis=1)
        df['其它属性9'] = df['其它属性9'].replace(".00","")

        df = df[['供应商商品编码','组合商品实体编码','其它属性8','组合装国标码','其它属性7','其它属性9','其它属性2','组合颜色规格']]
        df['来源文件'] = ""
        return df

    except Exception as e:
        print(f"处理文件时出错: {str(e)}")
        return None


# 执行函数
brand_mod = int(input('输入恢复模式\n0，全部品牌。 1，常规品牌：'))
filtered_data = select_and_filter_file(brand_mod)

# 显示筛选结果（可选）
if filtered_data is not None and not filtered_data.empty:
    filtered_data.to_excel(r"\\PCWCtz\童装\童装数据资料\temp_excel\1成品汇总.xlsx", index=False)
    print("可修改编码已更新到汇总表中")