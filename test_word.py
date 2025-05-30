import os
import sys
import getpass

# 添加项目路径到sys.path
project_root = r"c:\Users\zhouxingchen\UFO_Zuler"
if project_root not in sys.path:
    sys.path.append(project_root)

from ufo.automator.app_apis.word.wordclient import (
    WordWinCOMReceiver, 
    SaveAsCommand, 
    InsertTableCommand, 
    SelectTextCommand,
    SetFontCommand
)

def test_specific_saveas():
    """测试特定的SaveAs参数"""
    print("=== 测试特定SaveAs参数 ===")
    
    try:
        # Word的CLSID
        word_clsid = "Word.Application"
        
        # 创建Word接收器
        receiver = WordWinCOMReceiver(
            app_root_name="WINWORD.EXE",
            process_name="Word",
            clsid=word_clsid
        )
        
        print(f"客户端创建: {receiver.client is not None}")
        print(f"COM对象: {receiver.com_object is not None}")
        
        if receiver.com_object:
            print("✅ Word连接成功!")
            print(f"当前文档: {receiver.com_object.Name}")
            
            # 在文档中添加一些测试内容
            try:
                receiver.com_object.Range().Text = "这是SaveAs测试文档\n\n使用参数:\nfile_dir='C:\\Users\\%USERNAME%\\Desktop'\nfile_name='test'\nfile_ext='.docx'"
            except Exception as e:
                print(f"添加文档内容失败: {e}")
            
            # 测试您指定的SaveAs参数
            print(f"\n--- 测试指定的SaveAs参数 ---")
            print("参数: file_dir='C:\\Users\\%USERNAME%\\Desktop', file_name='test', file_ext='.docx'")
            
            save_command = SaveAsCommand(
                receiver=receiver,
                params={
                    "file_dir": "C:\\Users\\%USERNAME%\\Desktop",
                    "file_name": "test",
                    "file_ext": ".docx"
                }
            )
            
            try:
                result = save_command.execute()
                print(f"✅ SaveAs执行结果: {result}")
                
                # 验证文件是否真的被创建
                username = getpass.getuser()
                expected_file = f"C:\\Users\\{username}\\Desktop\\test.docx"
                
                if os.path.exists(expected_file):
                    print(f"✅ 文件成功创建: {expected_file}")
                    file_size = os.path.getsize(expected_file)
                    print(f"   文件大小: {file_size} 字节")
                else:
                    print(f"❌ 文件未找到: {expected_file}")
                    
            except Exception as e:
                print(f"❌ SaveAs执行失败: {e}")
                import traceback
                traceback.print_exc()
            
        else:
            print("❌ Word连接失败!")
            
    except Exception as e:
        print(f"测试过程中出错: {e}")
        import traceback
        traceback.print_exc()

def test_word_connection():
    """测试Word连接"""
    try:
        # Word的CLSID
        word_clsid = "Word.Application"
        
        # 创建Word接收器
        receiver = WordWinCOMReceiver(
            app_root_name="WINWORD.EXE",
            process_name="Word",  # 或者使用具体的文档名如"文档1.docx"
            clsid=word_clsid
        )
        
        print(f"客户端创建: {receiver.client is not None}")
        print(f"COM对象: {receiver.com_object is not None}")
        
        if receiver.com_object:
            print("✅ Word连接成功!")
            print(f"当前文档: {receiver.com_object.Name}")
            
            # 在文档中添加一些内容
            try:
                # 插入标题文本
                receiver.com_object.Range().Text = "Word自动化测试文档\n\n这是一个测试文档，用于验证Word COM功能。\n\n"
                
                # 移动到文档末尾
                end_range = receiver.com_object.Range()
                end_range.Collapse(0)  # 移动到末尾
                end_range.InsertAfter("下面将插入一个表格：\n")
                
            except Exception as e:
                print(f"添加文档内容失败: {e}")
            
            # 测试插入表格功能
            print("\n--- 测试插入表格 ---")
            insert_table_command = InsertTableCommand(
                receiver=receiver,
                params={
                    "rows": 3,
                    "columns": 4
                }
            )
            
            table_result = insert_table_command.execute()
            print(f"插入表格结果: {table_result}")
            
            # 测试文本选择功能
            print("\n--- 测试文本选择 ---")
            try:
                select_text_command = SelectTextCommand(
                    receiver=receiver,
                    params={
                        "text": "测试文档"
                    }
                )
                
                text_result = select_text_command.execute()
                print(f"文本选择结果: {text_result}")
            except Exception as e:
                print(f"文本操作失败: {e}")
            
            # 测试字体设置
            print("\n--- 测试字体设置 ---")
            try:
                # 先选择标题文本
                finder = receiver.com_object.Range().Find
                finder.Text = "Word自动化测试文档"
                if finder.Execute():
                    finder.Parent.Select()
                
                set_font_command = SetFontCommand(
                    receiver=receiver,
                    params={
                        "font_name": "Arial",
                        "font_size": 16
                    }
                )
                
                font_result = set_font_command.execute()
                print(f"字体设置结果: {font_result}")
            except Exception as e:
                print(f"字体设置失败: {e}")
            
            # 测试多种格式的保存功能
            print("\n--- 测试保存功能 ---")
            test_save_formats(receiver)
            
        else:
            print("❌ Word连接失败!")
            print("请确保Word已经启动，或者程序会自动创建新文档")
            
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()

def test_save_formats(receiver):
    """测试不同格式的保存功能"""
    
    # 使用桌面路径（带%USERNAME%变量）
    save_dir = "C:\\Users\\%USERNAME%\\Desktop\\word_test_output"
    
    print(f"保存目录: {save_dir}")
    
    # 测试保存为不同格式到桌面
    formats_to_test = [
        (".docx", "Word文档"),
        (".pdf", "PDF文件"),
        (".txt", "纯文本"),
        (".rtf", "RTF格式"),
        (".html", "网页格式")
    ]
    
    for file_ext, format_name in formats_to_test:
        print(f"\n测试保存为{format_name} ({file_ext})")
        
        save_command = SaveAsCommand(
            receiver=receiver,
            params={
                "file_dir": save_dir,
                "file_name": f"test_word_output_{file_ext[1:]}",  # 去掉点号
                "file_ext": file_ext
            }
        )
        
        try:
            result = save_command.execute()
            print(f"✅ {format_name}保存结果: {result}")
            
            # 验证文件是否真的被创建
            username = getpass.getuser()
            actual_dir = save_dir.replace("%USERNAME%", username)
            expected_file = os.path.join(actual_dir, f"test_word_output_{file_ext[1:]}{file_ext}")
            
            if os.path.exists(expected_file):
                print(f"   ✓ 文件已创建: {expected_file}")
            else:
                print(f"   ⚠ 文件未找到: {expected_file}")
                
        except Exception as e:
            print(f"❌ {format_name}保存失败: {e}")

def test_word_with_existing_document():
    """测试连接到已存在的Word文档"""
    try:
        # 如果您有特定的Word文档，可以指定文档名
        word_clsid = "Word.Application"
        
        receiver = WordWinCOMReceiver(
            app_root_name="WINWORD.EXE",
            process_name="文档1.docx",  # 替换为您的实际文档名
            clsid=word_clsid
        )
        
        print(f"连接特定文档:")
        print(f"客户端创建: {receiver.client is not None}")
        print(f"COM对象: {receiver.com_object is not None}")
        
        if receiver.com_object:
            print(f"✅ 成功连接到文档: {receiver.com_object.Name}")
            
            # 对现有文档进行保存测试
            print("\n--- 现有文档保存测试 ---")
            save_command = SaveAsCommand(
                receiver=receiver,
                params={
                    "file_dir": "C:\\Users\\%USERNAME%\\Desktop",  # 使用桌面路径
                    "file_name": "existing_doc_backup",
                    "file_ext": ".docx"
                }
            )
            
            result = save_command.execute()
            print(f"现有文档保存结果: {result}")
        else:
            print("❌ 未找到指定文档")
            
    except Exception as e:
        print(f"连接特定文档时出错: {e}")

def test_save_edge_cases(receiver):
    """测试保存功能的边界情况"""
    print("\n--- 测试保存边界情况 ---")
    
    # 测试空参数
    print("1. 测试默认参数保存")
    save_command = SaveAsCommand(
        receiver=receiver,
        params={}
    )
    result = save_command.execute()
    print(f"默认参数保存结果: {result}")
    
    # 测试无效路径
    print("\n2. 测试无效路径")
    save_command = SaveAsCommand(
        receiver=receiver,
        params={
            "file_dir": "C:\\invalid\\path\\that\\does\\not\\exist",
            "file_name": "test_invalid_path",
            "file_ext": ".docx"
        }
    )
    result = save_command.execute()
    print(f"无效路径保存结果: {result}")
    
    # 测试中文文件名到桌面
    print("\n3. 测试中文文件名到桌面")
    save_command = SaveAsCommand(
        receiver=receiver,
        params={
            "file_dir": "C:\\Users\\%USERNAME%\\Desktop",
            "file_name": "中文测试文档",
            "file_ext": ".docx"
        }
    )
    result = save_command.execute()
    print(f"中文文件名保存结果: {result}")

if __name__ == "__main__":
    # 首先测试您指定的特定SaveAs参数
    test_specific_saveas()
    
    # print("\n" + "="*60)
    # print("=== Word连接和SaveAs完整测试 ===")
    # test_word_connection()
    
    # print("\n" + "="*50)
    # print("=== 特定文档连接测试 ===")
    # test_word_with_existing_document()
    
    # 如果基本连接成功，再测试边界情况
    try:
        receiver = WordWinCOMReceiver(
            app_root_name="WINWORD.EXE",
            process_name="Word",
            clsid="Word.Application"
        )
        
        if receiver.com_object:
            print("\n" + "="*50)
            test_save_edge_cases(receiver)
        
    except Exception as e:
        print(f"边界测试初始化失败: {e}")
    
    print("\n=== 测试完成 ===")