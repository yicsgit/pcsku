from sys import exception

import pandas as pd
import re
import numpy as np
import pymysql
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
    print("----------------1件表格开始------------------------------------------------------------")
    try:
        df=pd.read_excel(path,header=0)# ,names=['供应商商品编码','商品编码'], dtype = {'商品编码': str} 在read里写names会多一行表头，可能是dtype强制字符导致的
        if df.empty:
            print('提交表格为空，已退出进程')
            return  # 或 sys.exit(1)
        df.columns = ['供应商商品编码','商品编码','其它属性8','国标码','其它属性7','吊牌价','其它属性2','颜色及尺码','来源文件']

        pd.set_option('future.no_silent_downcasting', True)
        # df = df.replace(r"\s",'',regex=True) # 替换任意空白字符
        # df = df.replace('替代码,要交',None,regex=True) # 替换掉“替代码,要交”

        df = df.replace({r'\s':'','临时占位,要交':None}, regex=True) # 字典或列表可同时批量替换
        # 应用转换函数
        df['商品编码'] = df['商品编码'].apply(convert_value)
        df = df.dropna(subset=['商品编码']) # 删除转换后为 None 的行（原始 NaN 和空字符串）
        max_row0 = len(df)

        # df = df.map(lambda x: None if pd.isna(x) else x) # 全表替换nan到None，有时候不成功，可能原因是没选定列
        df = df.replace(np.nan,None) # 全表替换NaN到None
        # 国标码验证
        # 将国标码转换为字符串
        df['国标码'] = df['国标码'].astype(str)
        # 定义一个函数来处理国标码
        def process_standard_code(code):
            # 检查前三位是否是069
            if code.startswith('069'):
                return '69' + code[3:]
            # 检查是否包含引号且前三位是'69
            elif code.startswith("'69"):
                return '69' + code[3:]
            # 其他情况保持不变
            return code
        # 应用函数处理国标码
        df['国标码'] = df['国标码'].apply(process_standard_code)

        df.iloc[:,0:2] = df.iloc[:,0:2].apply(lambda x: x.str.upper()) # 字符串转大写方法一
        # 字符串转大写方法二，加入if判断字符串，否则全表大写的同时会将价格数字也转化为字符串
        # df = df.apply map(lambda x: str(x).upper() if x.dtype == 'object' else x)
        df=df.drop_duplicates(subset="商品编码",inplace=False)# 删重
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
            blank_mask = clean_df.iloc[:,0].isna() #iloc 列的数字是左闭右开
            hanzi_mask = clean_df.iloc[:,0:2].apply(lambda col:col.astype(str).str.contains(r'[\u4e00-\u9fa5]',na=False)).any(axis=1)
            two_mask = clean_df.iloc[:,1].astype(str).str.contains(r'-[A-Z0-9]{4}-[A-Z0-9]', regex=True)
            combined_mask = error_mask | blank_mask | hanzi_mask | two_mask
            # print(combined_mask)
            return np.where(combined_mask)[0].tolist() if combined_mask.any() else False
            # where(mask)返回的是多元组：(([0, 2]),)   。加[0] 返回的是单元组：([0, 2])    。加[0].tolist 返回的是一维数组：[0, 2]
            # return clean_df.iloc[np.where(combined_mask)[0], 1].values # 这是返回的编码值的写法
        err_index = filter_err(df)
        if err_index is not False :
            err_sku = df.iloc[err_index, 1] # 获取有错的对应的编码
            df = df.drop(df.index[err_index])
            max_row2 = len(df)
            print(str(max_row1-max_row2) + "个错误编码，序号和编码如下：")
            print(err_sku)
        max_row2  = len(df)
        print("--资料校验完成--")
        df["衣服数量"]=df['供应商商品编码'].str.extract(r'-([A-Z0-9]+)-')
        df["衣服数量"]=df["衣服数量"].str.len()/2
        # pattern=re.compile(r"(2[0-9][TDGLR]|T)[A-Z][0-9A-Z][0-9]([A-SU-Z][0-9A-Z]\d{2}|T\d{3}){1,3}")
        # num_cloth=df['供应商商品编码'].str.split("-",expand=True)
        # num_cloth=num_cloth[1].str.len()
        # print(num_cloth)
        # one_cloth=num_cloth==2
        # df.loc[one_cloth,"供应商商品编码"]

        pattern=r"""((2[0-9][TDGLRK]|T)[A-Z][0-9A-Z][0-9])
        ([A-SU-Z][0-9A-Z]\d{2}|T\d{3}){1}
        ([A-SU-Z][0-9A-Z]\d{2}|T\d{3})?
        ([A-SU-Z][0-9A-Z]\d{2}|T\d{3})?
        ([A-SU-Z][0-9A-Z]\d{2}|T\d{3})?
        (((2[0-9][TDGLR]|T)[A-Z][0-9A-Z][0-9])([A-SU-Z][0-9A-Z]\d{2}|T\d{3})([A-SU-Z][0-9A-Z]\d{2}|T\d{3})?)?
        """

        regex_df=df['供应商商品编码'].str.extract(pattern,flags=re.VERBOSE)
        regex_df.columns=["款1","款1首字母","款1图1","款1图2","款1图3","款1图4","款2&图2","款2","款2首字母","款2图1","款2图2"]
        regex_df = regex_df.where(pd.notnull(regex_df), None) # 也可以用上面两种替换方法，此方式适合2.0版本
        df=pd.concat([df,regex_df],axis=1)
        # 剔除错误编码
        err_df=df.query("款1.isnull() or 衣服数量 not in [1,2]\
                        or (衣服数量==2 and 款2.isnull()) or (衣服数量==1 and 款2.notnull())")
        if len(err_df) !=0 :
            # df=df[~df.index.isin(err_df.index)]# 删掉不需要行 另一种写法
            df = df.drop(err_df.index)# 删掉不需要行
            max_row3 = len(df)
            print(str((max_row2-max_row3))+ "个错误编码，序号和编码如下：")
            print(err_df.iloc[:, 0:2])
            # err_df.to_excel("D:/0影刀导入文件夹/错误_字符问题.xlsx",index=False)
        df=df.drop(["款1首字母","款2&图2","款2首字母"],axis=1)# 删掉不需要列
        # 将dataframe 转换成 插入数据库的数据格式 的两种写法
        insert_df = [tuple(row) for _, row in df.iterrows()] # [(),(),()]
        # insert_df = df.values.tolist() # [[],[],[]]
        #---------------------------上传到mysql数据库中的product table------------------------
        print(str(len(insert_df))+'个正确编码(其中重复编码数量已删除)' )
        conn = pymysql.connect(host='localhost',user='root',password='123456',database='tz_sql')
        cursor = conn.cursor()

        try:
            table_name ="product_sku"
            cursor.execute("TRUNCATE TABLE " + table_name)  # 先清空数据表
            insert_query = 'INSERT INTO ' + table_name + ' VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'
            cursor.executemany(insert_query,insert_df)
            conn.commit()
            print("成功上传到Mysql\n")
        except Exception as e:
            print(f"错误:{e}")
            conn.rollback()

        cursor.close()
        conn.close()
    except Exception as e:
        print(f"错误{e}")






if __name__ == "__main__":
    excel_path = r"\\PCWCtz\童装\童装数据资料\temp_excel\1成品汇总.xlsx"
    clean_upload(excel_path)
