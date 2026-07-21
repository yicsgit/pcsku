import os
import time
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from tkinter import Tk, filedialog
import chardet
import xlrd
# from productsku_to_excel import clear_product

def merge_files_to_excel():
    # 选择文件夹
    root = Tk()
    root.withdraw() # 隐藏主窗口
    folder_path = filedialog.askdirectory(title="请选择包含数据文件的文件夹")

    if not folder_path:
        excel_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
        )
        if not excel_path:
            print("未选择文件")
            quit()
        try:
            file_name = os.path.basename(excel_path)
            df_title = ['内部T编码', '款色图', '成品编码', '款号', '国标码', '辅助码', '吊牌价', '品名', '颜色及尺码',
                        '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25',
                        '26', '27', '28', '29', '30']
            wb = load_workbook(excel_path)
            active_sheet = wb.active.title
            wb.close()
            df = pd.read_excel(excel_path, sheet_name=active_sheet, header=None, names=df_title)
            df['来源文件'] = file_name
            result_df = df
        except Exception as e:
            print(f"读取Excel文件 {file_name} 出错: {str(e)}")
    else:
        # 创建结果DataFrame
        result_df = pd.DataFrame()

        # 遍历文件夹中的文件
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            df_title = ['内部T编码','款色图','成品编码','款号','国标码','辅助码','吊牌价','品名','颜色及尺码','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30']
            # 处理Excel文件
            if file.endswith('.xlsx'):
                try:
                    wb = load_workbook(file_path)
                    active_sheet = wb.active.title
                    wb.close()
                    df = pd.read_excel(file_path,sheet_name=active_sheet,header=None,names=df_title)
                    df['来源文件'] = file
                    result_df = pd.concat([result_df, df], ignore_index=True)
                except Exception as e:
                    print(f"读取Excel文件 {file} 出错: {str(e)}")

            elif file.endswith('.xls'):
                try:
                    df = pd.read_excel(file_path,sheet_name=0,header=None,names=df_title,engine = 'xlrd')
                    df['来源文件'] = file
                    result_df = pd.concat([result_df, df], ignore_index=True)
                except Exception as e:
                    print(f"读取Excel文件 {file} 出错: {str(e)}")

            # 处理TXT文件
            elif file.endswith('.txt'):
                try:
                    with open(file_path, 'rb') as f:
                        raw_data = f.read(50000)  # 读取前5万字节用于检测编码类型，足够判断
                        编码类型 = chardet.detect(raw_data)
                    df = pd.read_csv(file_path, sep='\t', header=None,names=df_title,encoding=编码类型['encoding'])
                    df['来源文件'] = file
                    result_df = pd.concat([result_df, df], ignore_index=True)
                except Exception as e:
                    print(f"读取TXT文件 {file} 出错: {str(e)}")


    result_df["衣服数量"] = result_df['内部T编码'].str.extract(r'-([A-Z0-9]+)-')
    result_df["衣服数量"] = result_df["衣服数量"].str.len() / 2
    if (result_df["衣服数量"] == 4).any():
        print("请注意，有4件装编码")
    # 创建条件掩码。满足条件则件数=1，不满足的保持原值
    mask = result_df["内部T编码"].str.replace("-","") == result_df["成品编码"]
    result_df.loc[mask, '衣服数量'] = 1




    def filter_err(clean_df):
        error_mask = clean_df.apply(lambda col: col.astype(str).str.contains('错误|#N|#R', na=False)).any(axis=1)
        blank_mask = clean_df.iloc[:, 0].isna()  # iloc 列的数字是左闭右开
        hanzi_mask = clean_df.iloc[:, 0:2].apply(
            lambda col: col.astype(str).str.contains(r'[\u4e00-\u9fa5]', na=False)).any(axis=1)
        two_mask = clean_df.iloc[:, 1].astype(str).str.contains(r'-[A-Z0-9]{4}-[A-Z0-9]', regex=True)
        combined_mask = error_mask | blank_mask | hanzi_mask | two_mask
        # print(combined_mask)
        return np.where(combined_mask)[0].tolist() if combined_mask.any() else False
        # where(mask)返回的是多元组：(([0, 2]),)   。加[0] 返回的是单元组：([0, 2])    。加[0].tolist 返回的是一维数组：[0, 2]
        # return clean_df.iloc[np.where(combined_mask)[0], 1].values # 这是返回的编码值的写法


    product_df = result_df[result_df['衣服数量'] == 1]
    product_df = product_df[['内部T编码', '成品编码', '款号', '国标码', '辅助码', '吊牌价', '品名', '颜色及尺码', '来源文件']]
    err_index = filter_err(product_df)
    if err_index is not False:
        err_file_df = product_df.iloc[err_index,8]  # 获取有错的对应文件名
        err_file_name = err_file_df.unique()
        product_df = product_df.drop(product_df.index[err_index])
        print("单品有错误编码，来源文件如下：")
        print(err_file_name)
    # if len(product_df) != 0:
    product_df.to_excel(r"\\PCWCtz\童装\童装数据资料\temp_excel\1成品汇总.xlsx", index=False)




    def filter_combine_err(clean_df,num):
        error_mask = clean_df.apply(lambda col: col.astype(str).str.contains('错误|#N|#R', na=False)).any(axis=1)
        blank_mask = clean_df.iloc[:, 0:10].isna().any(axis=1) if num == 2 else clean_df.iloc[:, 0:14].isna().any(axis=1) # iloc 列的数字是左闭右开
        hanzi_mask = clean_df.iloc[:, 0].str.contains(r'[\u4e00-\u9fa5]', na=False)
        combined_mask = error_mask | blank_mask | hanzi_mask
        return np.where(combined_mask)[0].tolist() if combined_mask.any() else False
    #
    # def extract_2codes(text):
    #     pattern = r"((2[0-9][TDGLRO]|T)[A-Z][0-9A-Z][0-9])([A-SU-Z][0-9A-Z]\d{2}|T\d{3})([A-SU-Z][0-9A-Z]\d{2}|T\d{3})?"
    #     matches = re.findall(pattern, text)
    #     tup1 = matches[0]
    #     # 提取索引0,2,3的元素（跳过索引1）
    #     elements1 = [tup1[i] for i in [0, 2, 3] if i < len(tup1)]
    #     tup2 = matches[1]
    #     # 提取索引0,2,3的元素（跳过索引1）
    #     elements2 = [tup2[i] for i in [0, 2, 3] if i < len(tup2)]
    #     # 过滤空字符串并拼接
    #     return ''.join(e for e in elements1 if e != ''),''.join(e for e in elements2 if e != '')
        # return matches[0] if len(matches) > 0 else None, matches[1] if len(matches) > 1 else None
    suit2_df = result_df[result_df['衣服数量'] == 2]
    suit2_df = suit2_df[['内部T编码','来源文件','成品编码','款号','国标码','辅助码','吊牌价','品名','颜色及尺码',
                       '10','11','12','13','14']]
    suit2_df.columns = ['2件编码','来源文件','属性8_1','属性8_2','国标码_1','国标码_2','属性7_1','属性7_2','吊牌价_1','吊牌价_2','属性2_1','属性2_2','颜色_1','颜色_2']
    err2_index = filter_combine_err(suit2_df,2)
    if err2_index is not False:
        err2_file_df = suit2_df.iloc[err2_index,1]  # 获取有错的对应文件名
        err2_file_name = err2_file_df.unique()
        suit2_df = suit2_df.drop(suit2_df.index[err2_index])
        print("两件有错误编码，来源文件如下：")
        print(err2_file_name)

    # suit2_df[['正则1', '正则2']] = suit2_df['2件编码'].apply(lambda x: pd.Series(extract_2codes(x)))
    # # suit2_df['正则1'] = suit2_df['正则1'].apply(lambda s: s[0].join(filter(None, s[2:4])))
    # # suit2_df['正则2'] = suit2_df['正则2'].apply(lambda s: s[0].join(filter(None, s[2:4])))
    if len(suit2_df) != 0:
        suit2_df.to_excel(r"\\PCWCtz\童装\童装数据资料\temp_excel\2件套汇总.xlsx", index=False)

    suit3_df = result_df[result_df['衣服数量'] == 3]
    suit3_df = suit3_df[['内部T编码','来源文件','成品编码','款号','国标码','辅助码','吊牌价','品名','颜色及尺码',
                       '10','11','12','13','14','15','16','17','18','19','20']]
    suit3_df.columns = ['3件编码', '来源文件', '属性8_1', '属性8_2', '属性8_3', '国标码_1', '国标码_2', '国标码_3', '属性7_1', '属性7_2', '属性7_3',
                        '吊牌价_1', '吊牌价_2', '吊牌价_3', '属性2_1', '属性2_2', '属性2_3', '颜色_1', '颜色_2', '颜色_3']
    err3_index = filter_combine_err(suit3_df,3)
    if err3_index is not False:
        err3_file_df = suit3_df.iloc[err3_index,1]  # 获取有错的对应文件名
        err3_file_name = err3_file_df.unique()
        suit3_df = suit3_df.drop(suit3_df.index[err3_index])
        print("三件有错误编码，来源文件如下：")
        print(err3_file_name)
    if len(suit3_df) != 0:
        suit3_df.to_excel(r"\\PCWCtz\童装\童装数据资料\temp_excel\3件套汇总.xlsx", index=False)


    print(f"合并完成，成品和套装汇总已保存,{time.ctime()}")
    time.sleep(2)
    print("------merger-success-----------\n" * 3)

if __name__ == "__main__":
    merge_files_to_excel()