from download_product_from_mysql import download_to_excel as product_download_to_excel

from product_to_sql import clean_upload as product_clean_upload
from combine2_to_sql import clean_upload as combine2_clean_upload
from combine3_to_sql import clean_upload as combine3_clean_upload

import sys,os,time       
from threading import Timer


# 在Python脚本中添加（等待30秒自动关闭）
def close_window():
    print("\n[超时自动关闭]")
    os._exit(0)
try:
    if os.path.exists(r"\\PCWCtz\童装\童装数据资料\temp_excel\2件套汇总.xlsx"):
        excel_path = r"\\PCWCtz\童装\童装数据资料\temp_excel\2件套汇总.xlsx"
        combine2_clean_upload(excel_path)
        
    if os.path.exists(r"\\PCWCtz\童装\童装数据资料\temp_excel\3件套汇总.xlsx"):
        excel_path = r"\\PCWCtz\童装\童装数据资料\temp_excel\3件套汇总.xlsx"
        combine3_clean_upload(excel_path)

    excel_path = r"\\PCWCtz\童装\童装数据资料\temp_excel\1成品汇总.xlsx"
    product_clean_upload(excel_path)
    product_download_to_excel(1)
    print("\n----------所有表格完成--------------------------------------------------------------\n")

    timeout = int(sys.argv[1]) if len(sys.argv) > 1 else 30  # 默认30秒
    Timer(timeout, close_window).start()  # 启动倒计时
    print(f"------------------------{time.ctime()}--------------------------------------\n窗口将在{timeout}秒后自动关闭")
    input("按任意键键立即退出")  # 手动退出
except Exception as e:
    print(f"错误: {str(e)}")

finally:
    os._exit(0)
