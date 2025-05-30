# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import os
import time
import tempfile
from pathlib import Path
import getpass
from typing import Dict, Type

from ufo.automator.app_apis.basic import WinCOMCommand, WinCOMReceiverBasic
from ufo.automator.basic import CommandBasic


class WordWinCOMReceiver(WinCOMReceiverBasic):
    """
    The base class for Windows COM client.
    """

    _command_registry: Dict[str, Type[CommandBasic]] = {}

    def get_object_from_process_name(self) -> None:
        """
        Get the object from the process name.
        :return: The matched object.
        """
        try:
            print(f"DEBUG: Starting Word get_object_from_process_name")
            
            if not self.client:
                print("DEBUG: Word client is None, returning None")
                return None
            
            # 获取所有文档
            documents = self.client.Documents
            print(f"DEBUG: Documents collection count = {documents.Count}")
            
            object_name_list = [doc.Name for doc in documents]
            print(f"DEBUG: Documents collection = {object_name_list}")
            
            # 如果没有打开的文档，创建一个新文档
            # if documents.Count == 0:
            #     print("DEBUG: No documents open, creating a new document")
            #     try:
            #         new_doc = self.client.Documents.Add()
            #         print(f"DEBUG: New document created: {new_doc.Name}")
            #         return new_doc
            #     except Exception as e:
            #         print(f"DEBUG: Failed to create new document: {e}")
            #         return None
            
            # 如果有文档，尝试匹配
            matched_object = self.app_match(object_name_list)
            print(f"DEBUG: app_match result = {matched_object}")
            
            # 如果process_name只是"Word"或匹配失败，返回第一个文档
            if (self.process_name.lower() in ["word", "winword", "microsoft word"] or 
                not matched_object):
                if documents.Count > 0:
                    first_document = documents.Item(1)
                    print(f"DEBUG: Returning first document: {first_document.Name}")
                    return first_document
            
            # 尝试找到匹配的文档
            for doc in documents:
                if doc.Name == matched_object:
                    print(f"DEBUG: Found matched document: {doc.Name}")
                    return doc
            
            # 如果都失败了，返回第一个文档
            if documents.Count > 0:
                first_document = documents.Item(1)
                print(f"DEBUG: Fallback: returning first document: {first_document.Name}")
                return first_document
            
            print("DEBUG: No documents available and failed to create new one")
            return None
            
        except Exception as e:
            print(f"DEBUG: Exception in Word get_object_from_process_name: {e}")
            import traceback
            traceback.print_exc()
            return None

    def insert_table(self, rows: int, columns: int) -> object:
        """
        Insert a table at the end of the document.
        :param rows: The number of rows.
        :param columns: The number of columns.
        :return: The inserted table.
        """

        # Get the range at the end of the document
        end_range = self.com_object.Range()
        end_range.Collapse(0)  # Collapse the range to the end

        # Insert a paragraph break (optional)
        end_range.InsertParagraphAfter()
        table = self.com_object.Tables.Add(end_range, rows, columns)
        table.Borders.Enable = True

        return table

    def select_text(self, text: str) -> None:
        """
        Select the text in the document.
        :param text: The text to be selected.
        """
        finder = self.com_object.Range().Find
        finder.Text = text

        if finder.Execute():
            finder.Parent.Select()
            return f"Text {text} is selected."
        else:
            return f"Text {text} is not found."

    def select_paragraph(
        self, start_index: int, end_index: int, non_empty: bool = True
    ) -> None:
        """
        Select a paragraph in the document.
        :param start_index: The start index of the paragraph.
        :param end_index: The end index of the paragraph, if ==-1, select to the end of the document.
        :param non_empty: Whether to select the non-empty paragraphs only.
        """
        paragraphs = self.com_object.Paragraphs

        start_index = max(1, start_index)

        if non_empty:
            paragraphs = [p for p in paragraphs if p.Range.Text.strip()]

        para_start = paragraphs[start_index - 1].Range.Start

        # Select to the end of the document if end_index == -1
        if end_index == -1:
            para_end = self.com_object.Range().End
        else:
            para_end = paragraphs[end_index - 1].Range.End

        self.com_object.Range(para_start, para_end).Select()

    def select_table(self, number: int) -> None:
        """
        Select a table in the document.
        :param number: The number of the table.
        """
        tables = self.com_object.Tables
        if not number or number < 1 or number > tables.Count:
            return f"Table number {number} is out of range."

        tables(number).Select()
        return f"Table {number} is selected."

    def set_font(self, font_name: str = None, font_size: int = None) -> None:
        """
        Set the font of the selected text in the active Word document.

        :param font_name: The name of the font (e.g., "Arial", "Times New Roman", "宋体").
                        If None, the font name will not be changed.
        :param font_size: The font size (e.g., 12).
                        If None, the font size will not be changed.
        """
        selection = self.client.Selection

        if selection.Type == 0:  # wdNoSelection

            return "No text is selected to set the font."

        font = selection.Range.Font

        message = ""

        if font_name:
            font.Name = font_name
            message += f"Font is set to {font_name}."

        if font_size:
            font.Size = font_size
            message += f" Font size is set to {font_size}."

        print(message)
        return message

    def save_as(
        self, file_dir: str = "", file_name: str = "", file_ext: str = ""
    ) -> str:
        """
        Save the document to specified format with enhanced error handling.
        :param file_dir: The directory to save the file.
        :param file_name: The name of the file without extension.
        :param file_ext: The extension of the file.
        :return: Success message or error details.
        """

        word_ext_to_fileformat = {
            ".doc": 0,  # Word 97-2003 Document
            ".dot": 1,  # Word 97-2003 Template
            ".txt": 2,  # Plain Text (ASCII)
            ".rtf": 6,  # Rich Text Format (RTF)
            ".unicode.txt": 7,  # Unicode Text (custom extension, for clarity)
            ".htm": 8,  # Web Page (HTML)
            ".html": 8,  # Web Page (HTML)
            ".mht": 9,  # Single File Web Page (MHT)
            ".xml": 11,  # Word 2003 XML Document
            ".docx": 12,  # Word Document (default)
            ".docm": 13,  # Word Macro-Enabled Document
            ".dotx": 14,  # Word Template (no macros)
            ".dotm": 15,  # Word Macro-Enabled Template
            ".pdf": 17,  # PDF File
            ".xps": 18,  # XPS File
        }

        # Enhanced path handling
        if not file_dir:
            if hasattr(self.com_object, 'FullName') and self.com_object.FullName:
                file_dir = os.path.dirname(self.com_object.FullName)
            else:
                file_dir = self.get_safe_desktop_path()
        else:
            # Expand environment variables
            file_dir = os.path.expandvars(file_dir)
            # Expand user home directory
            file_dir = os.path.expanduser(file_dir)
            
        print(f"DEBUG: Resolved file_dir: {file_dir}")
            
        if not file_name:
            if hasattr(self.com_object, 'FullName') and self.com_object.FullName:
                file_name = os.path.splitext(os.path.basename(self.com_object.FullName))[0]
            else:
                file_name = f"word_document_{int(time.time())}"
            
        if not file_ext:
            file_ext = ".docx"  # Default to docx instead of pdf
        elif not file_ext.startswith('.'):
            file_ext = '.' + file_ext

        # Validate and create directory if needed
        try:
            Path(file_dir).mkdir(parents=True, exist_ok=True)
            print(f"DEBUG: Directory ensured: {file_dir}")
        except Exception as e:
            print(f"DEBUG: Cannot create directory {file_dir}: {e}")
            # Fallback to user's Documents folder
            file_dir = os.path.join(os.path.expanduser("~"), "Documents")
            try:
                Path(file_dir).mkdir(parents=True, exist_ok=True)
                print(f"DEBUG: Fallback to Documents: {file_dir}")
            except Exception:
                # Final fallback to temp directory
                file_dir = tempfile.gettempdir()
                print(f"DEBUG: Final fallback to temp: {file_dir}")

        # Generate unique filename if file exists
        base_file_path = os.path.join(file_dir, file_name + file_ext)
        file_path = base_file_path
        counter = 1
        
        while os.path.exists(file_path):
            name_with_counter = f"{file_name}_{counter}"
            file_path = os.path.join(file_dir, name_with_counter + file_ext)
            counter += 1

        if file_path != base_file_path:
            print(f"DEBUG: File exists, using unique name: {file_path}")

        # Get file format
        file_format = word_ext_to_fileformat.get(file_ext.lower(), 12)  # Default to docx

        try:
            # Check if com_object exists
            if self.com_object is None:
                return "Error: No active Word document found."
            
            # Attempt to save with different methods
            return self._attempt_save_with_fallback(file_path, file_format, file_ext)
            
        except Exception as e:
            return f"Failed to save the document to {file_path}. Error: {str(e)}"

    def _attempt_save_with_fallback(self, file_path: str, file_format: int, file_ext: str) -> str:
        """
        Attempt to save with multiple fallback strategies.
        """
        # Method 1: Standard SaveAs2 (Word 2013+)
        try:
            self.com_object.SaveAs2(file_path, FileFormat=file_format)
            if os.path.exists(file_path):
                file_size = os.path.getsize(file_path)
                return f"Document successfully saved to {file_path} (Size: {file_size} bytes)"
            else:
                return f"SaveAs2 completed but file not found at {file_path}"
        except Exception as e1:
            print(f"DEBUG: SaveAs2 failed: {e1}")
            
            # Method 2: Standard SaveAs (older Word versions)
            try:
                self.com_object.SaveAs(file_path, FileFormat=file_format)
                if os.path.exists(file_path):
                    file_size = os.path.getsize(file_path)
                    return f"Document successfully saved to {file_path} (Size: {file_size} bytes, legacy method)"
                else:
                    return f"SaveAs completed but file not found at {file_path}"
            except Exception as e2:
                print(f"DEBUG: Legacy SaveAs failed: {e2}")
                
                # Method 3: SaveAs without FileFormat parameter (let Word decide)
                try:
                    self.com_object.SaveAs2(file_path)
                    if os.path.exists(file_path):
                        return f"Document successfully saved to {file_path} (auto-format)"
                    else:
                        return f"Auto-format save completed but file not found"
                except Exception as e3:
                    print(f"DEBUG: Auto-format SaveAs failed: {e3}")
                
                # Method 4: Export method for PDF and XPS
                if file_ext.lower() in ['.pdf', '.xps']:
                    try:
                        export_format = 17 if file_ext.lower() == '.pdf' else 18  # wdExportFormatPDF, wdExportFormatXPS
                        self.com_object.ExportAsFixedFormat(
                            OutputFileName=file_path,
                            ExportFormat=export_format,
                            OpenAfterExport=False,
                            OptimizeFor=0,  # wdExportOptimizeForPrint
                            BitmapMissingFonts=True,
                            DocStructureTags=True,
                            CreateBookmarks=0,  # wdExportDocumentContent
                            UseISO19005_1=False
                        )
                        if os.path.exists(file_path):
                            return f"Document successfully exported to {file_path}"
                        else:
                            return f"Export completed but file not found"
                    except Exception as e4:
                        print(f"DEBUG: Export method failed: {e4}")
                
                # Method 5: Copy and save approach
                try:
                    # Create a new document and copy content
                    new_doc = self.client.Documents.Add()
                    
                    # Copy all content from original document
                    self.com_object.Range().Copy()
                    new_doc.Range().Paste()
                    
                    # Save the new document
                    new_doc.SaveAs2(file_path, FileFormat=file_format)
                    new_doc.Close()
                    
                    if os.path.exists(file_path):
                        return f"Document successfully saved to {file_path} (copy method)"
                    else:
                        return f"Copy method completed but file not found"
                        
                except Exception as e5:
                    print(f"DEBUG: Copy method failed: {e5}")
                
                # Final fallback: save as default Word format in Documents
                try:
                    fallback_path = os.path.join(
                        os.path.expanduser("~"), "Documents", 
                        f"UFO_Word_Export_{int(time.time())}.docx"
                    )
                    self.com_object.SaveAs2(fallback_path, FileFormat=12)
                    return f"Document saved to fallback location: {fallback_path}"
                except Exception as e6:
                    return f"All save methods failed. Last error: {str(e6)}"

    def get_safe_desktop_path(self) -> str:
        """
        Get a safe desktop path for saving files.
        """
        try:
            # Try multiple methods to get desktop path
            desktop_paths = [
                os.path.join(os.path.expanduser("~"), "Desktop"),
                os.path.join(os.path.expandvars("%USERPROFILE%"), "Desktop"),
                os.path.join(os.path.expandvars("%HOMEDRIVE%"), os.path.expandvars("%HOMEPATH%"), "Desktop"),
                os.path.join(os.path.expanduser("~"), "桌面"),  # 中文系统
            ]
            
            for path in desktop_paths:
                if os.path.exists(path) and os.access(path, os.W_OK):
                    print(f"DEBUG: Found desktop at: {path}")
                    return path
            
            # Fallback to Documents
            docs_path = os.path.join(os.path.expanduser("~"), "Documents")
            if os.path.exists(docs_path):
                print(f"DEBUG: Fallback to Documents: {docs_path}")
                return docs_path
            
            # Final fallback to user home
            return os.path.expanduser("~")
            
        except Exception as e:
            print(f"DEBUG: Error getting desktop path: {e}")
            # Final fallback to temp directory
            return tempfile.gettempdir()

    @property
    def type_name(self):
        return "COM/WORD"

    @property
    def xml_format_code(self) -> int:
        return 11


@WordWinCOMReceiver.register
class InsertTableCommand(WinCOMCommand):
    """
    The command to insert a table.
    """

    def execute(self):
        """
        Execute the command to insert a table.
        :return: The inserted table.
        """
        return self.receiver.insert_table(
            self.params.get("rows"), self.params.get("columns")
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "insert_table"


@WordWinCOMReceiver.register
class SelectTextCommand(WinCOMCommand):
    """
    The command to select text.
    """

    def execute(self):
        """
        Execute the command to select text.
        :return: The selected text.
        """
        return self.receiver.select_text(self.params.get("text"))

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "select_text"


@WordWinCOMReceiver.register
class SelectTableCommand(WinCOMCommand):
    """
    The command to select a table.
    """

    def execute(self):
        """
        Execute the command to select a table in the document.
        :return: The selected table.
        """
        return self.receiver.select_table(self.params.get("number"))

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "select_table"


@WordWinCOMReceiver.register
class SelectParagraphCommand(WinCOMCommand):
    """
    The command to select a paragraph.
    """

    def execute(self):
        """
        Execute the command to select a paragraph in the document.
        :return: The selected paragraph.
        """
        return self.receiver.select_paragraph(
            self.params.get("start_index"),
            self.params.get("end_index"),
            self.params.get("non_empty"),
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "select_paragraph"


@WordWinCOMReceiver.register
class SaveAsCommand(WinCOMCommand):
    """
    The command to save the document to a specific format.
    """

    def execute(self):
        """
        Execute the command to save the document to a specific format.
        :return: The result of saving the document.
        """
        # 首先检查com_object是否为None
        if self.receiver.com_object is None:
            return "Error: No active Word document found. Please ensure Word is running with an open document."
        
        # Get parameters with enhanced defaults
        file_dir = self.params.get("file_dir", "")
        file_name = self.params.get("file_name", "")
        file_ext = self.params.get("file_ext", ".docx")
        
        # Handle special path cases
        if file_dir and "%USERNAME%" in file_dir:
            # Replace %USERNAME% with actual username
            username = getpass.getuser()
            file_dir = file_dir.replace("%USERNAME%", username)
            print(f"DEBUG: Resolved %USERNAME% to: {file_dir}")
        
        # Ensure file extension starts with dot
        if file_ext and not file_ext.startswith('.'):
            file_ext = '.' + file_ext
            
        # Use safe desktop path if specified but invalid
        if file_dir and "Desktop" in file_dir and not os.path.exists(file_dir):
            file_dir = self.receiver.get_safe_desktop_path()
            print(f"DEBUG: Desktop fallback: {file_dir}")
        
        try:
            result = self.receiver.save_as(
                file_dir=file_dir,
                file_name=file_name,
                file_ext=file_ext,
            )
            return result
        except Exception as e:
            return f"SaveAs command failed: {str(e)}"

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "save_as"


@WordWinCOMReceiver.register
class SetFontCommand(WinCOMCommand):
    """
    The command to set the font of the selected text.
    """

    def execute(self):
        """
        Execute the command to set the font of the selected text.
        :return: The message of the font setting.
        """
        return self.receiver.set_font(
            self.params.get("font_name"), self.params.get("font_size")
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "set_font"
