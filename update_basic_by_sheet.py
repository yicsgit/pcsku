import pandas as pd
import numpy as np
import pymysql
# ------------------------------------------------------------------------------
def update_to_sql(update_mod,sht_name,table_name,values_type):
    insert_df = pd.read_excel(r"\\PCWCtz\童装\童装数据资料\童装通用模板.xlsx", header=0, sheet_name=sht_name)
    if update_mod == 1 : # 款式表格处理
        insert_df = insert_df.loc[:,["款号","商品全称（不带童）","简称","分类","验证分类列","品类","品牌","执行标准","安全类别（GB后面不带空格）","成分","季节","功能吊牌","功能性执行标准"]]
        insert_df["商品全称（不带童）"] = insert_df["商品全称（不带童）"].replace('^童','',regex=True) # 去掉第一个童
        insert_df["安全类别（GB后面不带空格）"] = insert_df["安全类别（GB后面不带空格）"].str.replace("GB +","GB",regex=True) # 去掉GB后的空格
    elif update_mod == 2 : # 印花表格处理
        trans_table = str.maketrans({'；':';','（':'(','）':')'}) #创建字符映射的转换表，配合translate方法使用
        insert_df['九宫格信息'] = insert_df['九宫格信息'].str.translate(trans_table)
    elif update_mod == 3 : # 白坯编码表格处理
        insert_df = insert_df.drop_duplicates(subset=['白坯的英文编码'])
        insert_df['白坯的中文编码']=insert_df['白坯的中文编码'].replace({'': None, np.nan: None})
        insert_df = insert_df.loc[:,['白坯的英文编码','白坯的中文编码']]
    elif update_mod == 4 : # 颜色表格处理
        insert_df = insert_df.loc[:,['颜色编码','颜色名']]
        insert_df = insert_df.dropna(subset=['颜色编码'])
    else :
        raise SystemExit("选择更新的模式错误，程序已结束")

    insert_df = insert_df.where(pd.notnull(insert_df), None)
    insert_df = [tuple(row) for _, row in insert_df.iterrows()]
                 # -----------------------------------------------
    conn = pymysql.connect(host='localhost',user='root',password='123456',database='tz_sql') # charset='utf8'
    cursor = conn.cursor()

    try:
        cursor.execute("TRUNCATE TABLE " + table_name)# 先清空数据表
        insert_sql = 'INSERT INTO ' + table_name + ' VALUES ' + values_type
        # insert_sql = f'''INSERT INTO {table_name} 30
        
        #                 VALUES {values_type}
        #                 ON DUPLICATE KEY UPDATE 
        #                 图名=VALUES(图名),
        #                 聚水潭大品牌=VALUES(聚水潭大品牌),
        #                 规格=VALUES(规格),
        #                 九宫格信息=VALUES(九宫格信息),
        #                 表格小品牌=VALUES(表格小品牌),
        #                 工艺=VALUES(工艺);'''
        cursor.executemany(insert_sql,insert_df) # 再新增数据内容
        print(sht_name + " 表格,更新完成")
        conn.commit()
    except Exception as e:
        print(f"错误:{e}")
        conn.rollback()

    cursor.close()
    conn.close()
                 # -----------------------------------------------

# ------------------------------------------------------------------------------
dic = { 1:['款式','basic_spu','(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'],
        2:['图','basic_pic','(%s,%s,%s,%s,%s,%s,%s)'],
        3:['白坯编码','basic_clothing_code','(%s,%s)'],
        4:['颜色','basic_color','(%s,%s)']
        }
a= int(input( '0:ALL,1:款式,2:图,3:白坯编码,4:颜色,5:款+白坯\n请输入你需要更新的表：'))
#开启计时
import time
start_time = time.time()
if a in [1,2,3,4] :
    update_to_sql(a,dic.get(a)[0],dic.get(a)[1],dic.get(a)[2])
elif a == 0 :
    for i in [1,2,3,4]:
        update_to_sql(i, dic.get(i)[0], dic.get(i)[1], dic.get(i)[2])
elif a == 5 :
    update_to_sql(1,dic.get(1)[0],dic.get(1)[1],dic.get(1)[2])
    update_to_sql(3,dic.get(3)[0],dic.get(3)[1],dic.get(3)[2])
else :
    print('输入数字有错误，请重试')
# 结束计时
end_time = time.time()
print(f"更新耗时：{end_time - start_time}秒")