import pymysql
import pandas as pd
import datetime
#---------------------------执行指定的sql文件，将成品表和组合表导出来---------------------
#-----------------------------------------------------------------------------------
# with open('C:/Users/Administrator/Documents/Navicat/MySQL/Servers/PC_sql/tz_sql/成品view.sql', 'r', encoding='utf-8') as f:
#     view_query=f.read()-- # cursor.execute(view_query)

def download_to_excel(view_mod):
    if view_mod==1:
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\view-单件成品.sql', 'r',encoding='utf-8') as f:
            view_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\成品-单件导出表.sql', 'r', encoding='utf-8') as f:
            cp_query=f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\组合-单件导出表.sql', 'r',encoding='utf-8') as f:
            zh_query = f.read()
    elif view_mod==2:
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\view-不改款且忽略印花下架.sql', 'r',encoding='utf-8') as f:
            view_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\成品-不改款且忽略印花下架.sql', 'r',encoding='utf-8') as f:
            cp_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\组合-不改款且忽略印花下架.sql', 'r',encoding='utf-8') as f:
            zh_query = f.read()
    elif view_mod==3:
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\view-越南公版.sql', 'r',encoding='utf-8') as f:
            view_query = f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\成品-单件导出表.sql', 'r', encoding='utf-8') as f:
            cp_query=f.read()
        with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\组合-越南公版.sql', 'r',encoding='utf-8') as f:
            zh_query = f.read()
    else:quit()
    with open(r'C:\Users\Administrator\Documents\Navicat\MySQL\Servers\PC_child\tz_sql\错误-单件成品.sql', 'r', encoding='utf-8') as f:
        err_query=f.read()

    with pymysql.connect(host='localhost',user='root',password='123456',database='tz_sql') as conn:
        with conn.cursor() as cursor:
            try:
                cursor.execute(view_query)
                print("成功更新单件视图view")
                cursor.execute(err_query)
                tuples_on_row = cursor.fetchall()
                if len(tuples_on_row) != 0:
                    print("有"+str(len(tuples_on_row))+"个错误，请检查")
                    columns = [i[0] for i in cursor.description] #.description 是结果的数据类型等信息
                    df = pd.DataFrame(tuples_on_row,columns=columns) # ["需要输入标题","20个title"]
                    df.to_excel(r'\\PCWCtz\童装\童装数据资料\temp_excel\错误_单件编码问题.xlsx', index=False)
            except Exception as e:
                print(f"错误:{e}")

            try:
                cursor.execute(cp_query)
                tuples_on_row = cursor.fetchall()
                if len(tuples_on_row) == 0:
                    print('没有正确的编码')
                else:
                    columns = [i[0] for i in cursor.description] #.description 是结果的数据类型等信息
                    df = pd.DataFrame(tuples_on_row,columns=columns) # [如果不用description则需要输入标题,"20个title"]
                    brand = df.loc[0,"品牌"]
                    current_time = datetime.datetime.now().strftime("%m%d%H%M%S")
                    df.to_excel(f'//PCWCtz/童装/童装数据资料/temp_excel/成品/{brand}{current_time}-成品导入.xlsx', index=False)
                    print("成品表下载完成",end="   ")

                    cursor.execute(zh_query)
                    tuples_on_row = cursor.fetchall()
                    columns = [i[0] for i in cursor.description]  # .description 是结果的数据类型等信息
                    df = pd.DataFrame(tuples_on_row, columns=columns)  # ["需要输入标题","20个title"]
                    df.to_excel(f'//PCWCtz/童装/童装数据资料/temp_excel/组合/{brand}{current_time}-组合导入.xlsx', index=False)
                    print("组合表下载完成")
                    cursor.execute('drop VIEW if exists temp_view;')
                    print("视图已清空")
            except Exception as e:
                print(f"错误:{e}")
    print(f"----------------1件表格结束-------------------------------------------------------\n\n")       
if __name__ == "__main__":
    a = int(input("1，正常模式，2，忽略印花下架且不改款，3，越南公版\n请输入模式："))
    download_to_excel(a)
    import sys,os   
    from threading import Timer
    def close_window():
        print("\n[超时自动关闭]")
        os._exit(0)

    try:
        timeout = int(sys.argv[1]) if len(sys.argv) > 1 else 30  # 默认30秒
        Timer(timeout, close_window).start()  # 启动倒计时
        print(f"--------------------------------------------------------------\n窗口将在{timeout}秒后自动关闭")
        input("按任意键键立即退出")  # 手动退出
    except Exception as e:
        print(f"错误: {str(e)}")

    finally:
        os._exit(0)