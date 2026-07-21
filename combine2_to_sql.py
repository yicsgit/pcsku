from sys import exception

import pandas as pd
import re
import numpy as np
import pymysql

import download_combine_from_mysql
#---------------------------读取成品编码，汇总清洗--------------------------------------
# 使用 apply() 精确处理不同类型的编码
def convert_value(x):
    if pd.isna(x):
        return None  # 标记为删除
    elif isinstance(x, (int, float)):
        # 将数字转为字符串（不带小数点）
        return str(int(x))
    elif isinstance(x, str):
        stripped = x.strip()
        return stripped if stripped != "" else None  # 标记空字符串为删除
    else:
        return str(x).strip()  # 其他类型转为字符串

def clean_upload(path):
    try:
        print("----------------2件表格开始------------------------------------------------------------")
        df=pd.read_excel(path,header=0)# ,names=['供应商商品编码','商品编码'], dtype = {'商品编码': str} 在read里写names会多一行表头，可能是dtype强制字符导致的
        if df.empty:
            print('提交表格为空，已退出当前函数')
            return # 返回None表示函数执行失败

        df.columns = ['2件编码','来源文件','属性8_1','属性8_2','国标码_1','国标码_2','属性7_1','属性7_2','吊牌价_1','吊牌价_2','属性2_1','属性2_2','颜色_1','颜色_2']

        pd.set_option('future.no_silent_downcasting', True)
        # df = df.replace(r"\s",'',regex=True) # 替换任意空白字符
        # df = df.replace('替代码,要交',None,regex=True) # 替换掉“替代码,要交”
        df = df.replace(r'\s','', regex=True) # 字典或列表可同时批量替换
        # 应用转换函数
        df['2件编码'] = df['2件编码'].apply(convert_value)
        df = df.dropna(subset=['2件编码']) # 删除转换后为 None 的行（原始 NaN 和空字符串）
        max_row0 = len(df)

        # 待添加的剔除条件：错误国标码的验证，替换为空可能有问题
        # df = df.map(lambda x: None if pd.isna(x) else x) # 全表替换nan到None，有时候不成功，可能原因是没选定列
        df = df.replace(np.nan,None) # 全表替换NaN到None
        # 国标码待验证
        # df['国标码'] = df['国标码'].apply(convert_value)
        # print(df['国标码'])
        # 使用正确的字符串大写方法，添加类型检查以避免非字符串值报错
        df.iloc[:,0] = df.iloc[:,0].apply(lambda x: x.upper() if isinstance(x, str) else x) # 字符串转大写方法一
        # 字符串转大写方法二，加入if判断字符串，否则全表大写的同时会将价格数字也转化为字符串
        # df = df.apply map(lambda x: str(x).upper() if x.dtype == 'object' else x)
        df=df.drop_duplicates(subset="2件编码",inplace=False)# 删重
        max_row1 = len(df)
        # print("----------------2.1-----------------------------------------------------------")
        # def filter_err(row): # 按行挑出错误编码
        #     if any((row.astype(str).str.contains('错误')) | (row.astype(str).str.contains('#N')) | (row.astype(str).str.contains('#R'))):
        #         return row.iloc[1] # 挑出行上有 错误,#NAME,#NA,#REF 字符的编码
        #     if row.iloc[0:1].astype(str).str.contains(r'[\u4e00-\u9fa5]').any():
        #         return row.iloc[1] # 挑出T编码和商品编码 有包含汉字的编码
        #     if re.search(r'-[A-Z0-9]{4}-[A-Z0-9]',row.iloc[1]):
        #         return row.iloc[1] # 挑出商品编码是两件装的编码
        #     return None
        # err_result = df.apply(filter_err, axis=1).dropna() # 轴为1，按行查询
        # df = df.drop(err_result.index)
        print("--编码转化完成--",end="   ")

        # 待添加的剔除条件：错误国标码的验证，替换掉“替代码,要交”，替换为空可能有问题
        def filter_err(clean_df) :
            error_mask = clean_df.apply(lambda col:col.astype(str).str.contains('错误|#N|#R',na=False)).any(axis=1)
            blank_mask = clean_df.iloc[:,0:10].isna().any(axis=1) #iloc 列的数字是左闭右开
            hanzi_mask = clean_df.iloc[:,0].str.contains(r'[\u4e00-\u9fa5]',na=False)
            combined_mask = error_mask | blank_mask | hanzi_mask
            # print(combined_mask)
            return np.where(combined_mask)[0].tolist() if combined_mask.any() else False
            # where(mask)返回的是多元组：(([0, 2]),)   。加[0] 返回的是单元组：([0, 2])    。加[0].tolist 返回的是一维数组：[0, 2]
            # return clean_df.iloc[np.where(combined_mask)[0], 1].values # 这是返回的编码值的写法
        err_index = filter_err(df)
        if err_index and err_index is not False:  # 确保有错误索引且不是False
            err_filenames = df.iloc[err_index, 1] # 获取有错的对应的文件名
            df = df.drop(df.index[err_index])
            err_filename = err_filenames.unique()
            print("两件有错误编码的文件如下：")
            # 确保err_filename是可打印的格式
            if len(err_filename) > 0:
                for filename in err_filename:
                    print(filename)
        max_row2  = len(df)
        # df = df.replace({'替代码,要交':None,'临时占位,要交':None}, regex=True) # 字典或列表可同时批量替换
        print("--资料校验完成--")

        def extract_2codes(text):
            pattern = r"((2[0-9][TDGLROK]|T)[A-Z][0-9A-Z][0-9])([A-SU-Z][0-9A-Z]\d{2}|T\d{3})([A-SU-Z][0-9A-Z]\d{2}|T\d{3})?([A-SU-Z][0-9A-Z]\d{2}|T\d{3})?"
            matches = re.findall(pattern, text)
            tup1 = matches[0]
            # 提取索引0,2,3的元素（跳过索引1）
            elements1 = [tup1[i] for i in [0, 2, 3,4] if i < len(tup1)]
            tup2 = matches[1]
            # 提取索引0,2,3的元素（跳过索引1）
            elements2 = [tup2[i] for i in [0, 2, 3,4] if i < len(tup2)]
            # 过滤空字符串并拼接
            return ''.join(e for e in elements1 if e != ''),''.join(e for e in elements2 if e != '')

        df[['正则1', '正则2']] = df['2件编码'].apply(lambda x: pd.Series(extract_2codes(x)))
        # 将dataframe 转换成 插入数据库的数据格式 的两种写法
        insert_df = [tuple(row) for _, row in df.iterrows()] # [(),(),()]
        # insert_df = df.values.tolist() # [[],[],[]]
        #---------------------------上传到mysql数据库中的product table------------------------
        print('两件有' + str(len(insert_df))+'个正确编码(其中重复编码数量已删除)',end='  ')
        conn = pymysql.connect(host='localhost',user='root',password='123456',database='tz_sql')
        cursor = conn.cursor()

        try:
            table_name ="sku_combine2"
            cursor.execute("TRUNCATE TABLE " + table_name)  # 先清空数据表
            insert_query = 'INSERT INTO ' + table_name + ' VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
            cursor.executemany(insert_query,insert_df)
            conn.commit()
            print("成功上传\n")
        except Exception as e:
            print(f"错误:{e}")
            conn.rollback()

        cursor.close()
        conn.close()
    except Exception as e:
        print(f"错误{e}")
    

    download_combine_from_mysql.download_to_excel(2)

if __name__ == "__main__":
    excel_path = r"\\PCWCtz\童装\童装数据资料\temp_excel\2件套汇总.xlsx"
    clean_upload(excel_path)
