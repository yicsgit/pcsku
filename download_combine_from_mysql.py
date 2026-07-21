import pymysql
import pandas as pd
import datetime
import warnings,os
warnings.filterwarnings('ignore', category=UserWarning)
#---------------------------执行指定的sql文件，将成品表和组合表导出来---------------------
#-----------------------------------------------------------------------------------
# with open('C:/Users/Administrator/Documents/Navicat/MySQL/Servers/PC_sql/tz_sql/成品view.sql', 'r', encoding='utf-8') as f:
#     view_query=f.read()-- # cursor.execute(view_query)

def download_to_excel(num):
    if num==2:
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\view-两件组合.sql', 'r',encoding='utf-8') as f:
            view_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\错误-两件组合.sql', 'r', encoding='utf-8') as f:
            err_query=f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\组合-两件导出表.sql', 'r',encoding='utf-8') as f:
            zh_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\拆分-第一件.sql', 'r',encoding='utf-8') as f:
            cf1_query = f.read()       
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\拆分-第二件.sql', 'r',encoding='utf-8') as f:
            cf2_query = f.read()
    elif num==3:
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\view-三件组合.sql', 'r',encoding='utf-8') as f:
            view_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\错误-三件组合.sql', 'r', encoding='utf-8') as f:
            err_query=f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\组合-三件导出表.sql', 'r',encoding='utf-8') as f:
            zh_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\拆分-第一件.sql', 'r',encoding='utf-8') as f:
            cf1_query = f.read()       
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\拆分-第二件.sql', 'r',encoding='utf-8') as f:
            cf2_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\拆分-第三件.sql', 'r',encoding='utf-8') as f:
            cf3_query = f.read()
    else:quit()


    # 初始化DataFrame变量，避免作用域问题
    cf_df1 = pd.DataFrame()
    cf_df2 = pd.DataFrame()
    cf_df3 = pd.DataFrame()
    
    with pymysql.connect(host='localhost',user='root',password='123456',database='tz_sql') as conn:
        with conn.cursor() as cursor:
            try:
                cursor.execute(view_query)
                print(f"成功更新{num}件视图view",end='  ')
                
                # 在try块内为DataFrame赋值
                cf_df1 = pd.read_sql(cf1_query, conn)
                cf_df2 = pd.read_sql(cf2_query, conn)
                if num==3:
                    cf_df3 = pd.read_sql(cf3_query, conn)

                cursor.execute(err_query)
                tuples_on_row = cursor.fetchall()
                if len(tuples_on_row) != 0:
                    print("有"+str(len(tuples_on_row))+"个错误，请检查")
                    columns = [i[0] for i in cursor.description] #.description 是结果的数据类型等信息
                    df = pd.DataFrame(tuples_on_row,columns=columns) # ["需要输入标题","20个title"]
                    df.to_excel(rf'\\PCWCtz\童装\童装数据资料\temp_excel\错误_{num}件编码问题.xlsx', index=False)
            except Exception as e:
                print(f"错误:{e}")

            try:
                cursor.execute(zh_query)
                tuples_on_row = cursor.fetchall()
                if len(tuples_on_row) == 0:
                    print(f'没有正确的{num}件编码')
                else:
                    columns = [i[0] for i in cursor.description]  # .description 是结果的数据类型等信息
                    df = pd.DataFrame(tuples_on_row, columns=columns)  # ["需要输入标题","20个title"]
                    current_time = datetime.datetime.now().strftime("%m%d%H%M%S")
                    df.to_excel(f'//PCWCtz/童装/童装数据资料/temp_excel/组合/{current_time}-{num}件组合导入.xlsx', index=False)
                    print(f"{num}件组合表下载完成",end="   ")
                    
                    cursor.execute('drop VIEW if exists temp_combine_view;')
                    print("组合视图已清空")
            except Exception as e:
                print(f"错误:{e}")
      
    # 读取现有Excel文件并重置索引，确保没有重复索引
    existing_df = pd.read_excel(r'\\PCWCtz\\童装\\童装数据资料\\temp_excel\\1成品汇总.xlsx')
    existing_df = existing_df.reset_index(drop=True)  # 重置索引，丢弃旧索引

    # 确保其他DataFrame也有唯一索引
    cf_df1 = cf_df1.reset_index(drop=True)
    cf_df2 = cf_df2.reset_index(drop=True)
    if num==3:
        cf_df3 = cf_df3.reset_index(drop=True)

    # 合并DataFrame
    combined_df = pd.concat([existing_df, cf_df1, cf_df2], ignore_index=True)
    if num==3:
        combined_df = pd.concat([combined_df, cf_df3], ignore_index=True)

    # 写入Excel文件
    combined_df.to_excel(r'\\PCWCtz\\童装\\童装数据资料\\temp_excel\\1成品汇总.xlsx', index=False, sheet_name='Sheet1')
    print("拆分单品插入到表完成")
    print(f"----------------{num}件表格结束-------------------------------------------------------\n\n")
    os.remove(fr"\\PCWCtz\童装\童装数据资料\temp_excel\{num}件套汇总.xlsx")
    # with pd.ExcelWriter(r'\\PCWCtz\童装\童装数据资料\temp_excel\1成品汇总.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    #     combined_df.to_excel(writer, index=False, sheet_name='Sheet1')    
if __name__ == "__main__":
    a = int(input("1,取消，2，两件组合，3，3件组合\n请输入模式："))
    if a==1:
        quit()
    download_to_excel(a)