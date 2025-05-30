import os
import sys

# 添加项目路径到sys.path
project_root = r"c:\Users\zhouxingchen\UFO_Zuler"
if project_root not in sys.path:
    sys.path.append(project_root)

from ufo.automator.app_apis.excel.excelclient import ExcelWinCOMReceiver, SaveAsCommand

def test_excel_connection():
    """测试Excel连接"""
    try:
        # Excel的CLSID
        excel_clsid = "Excel.Application"
        
        # 创建Excel接收器
        receiver = ExcelWinCOMReceiver(
            app_root_name="EXCEL.EXE",
            process_name="工作簿1.xlsx",  # 替换为您的Excel文件名
            clsid=excel_clsid
        )
        
        print(f"客户端创建: {receiver.client is not None}")
        print(f"COM对象: {receiver.com_object is not None}")
        
        if receiver.com_object:
            print("✅ Excel连接成功!")
            
            # 测试保存功能
            save_command = SaveAsCommand(
                receiver=receiver,
                params={
                    "file_dir": "",  # 空表示使用当前目录
                    "file_name": "test_output",
                    "file_ext": ".xlsx"
                }
            )
            
            result = save_command.execute()
            print(f"保存结果: {result}")
            
        else:
            print("❌ Excel连接失败!")
            
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_excel_connection()