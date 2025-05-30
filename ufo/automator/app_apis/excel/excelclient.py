# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import os
from typing import Any, Dict, List, Type, Union
import os
import tempfile
from pathlib import Path

import pandas as pd

from ufo.automator.app_apis.basic import WinCOMCommand, WinCOMReceiverBasic
from ufo.automator.basic import CommandBasic


class ExcelWinCOMReceiver(WinCOMReceiverBasic):
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
            print(f"DEBUG: Starting get_object_from_process_name")
            print(f"DEBUG: self.client = {self.client}")
            
            if not self.client:
                print("DEBUG: Client is None, returning None")
                return None
            
            # 获取所有工作簿
            workbooks = self.client.Workbooks
            print(f"DEBUG: Workbooks collection = {workbooks}")
            print(f"DEBUG: Workbooks count = {workbooks.Count}")
            
            object_name_list = []
            for i in range(1, workbooks.Count + 1):
                try:
                    workbook = workbooks.Item(i)
                    name = workbook.Name
                    object_name_list.append(name)
                    print(f"DEBUG: Workbook {i}: {name}")
                    print(f"DEBUG: Workbook {i} encoding: {name.encode('utf-8')}")
                except Exception as e:
                    print(f"DEBUG: Error getting workbook {i}: {e}")
            
            print(f"DEBUG: object_name_list = {object_name_list}")
            print(f"DEBUG: process_name = {self.process_name}")
            print(f"DEBUG: process_name encoding: {self.process_name.encode('utf-8')}")
            
            if not object_name_list:
                print("DEBUG: No workbooks found")
                return None
            
            # 尝试直接匹配
            for name in object_name_list:
                if name == self.process_name:
                    print(f"DEBUG: Direct match found: {name}")
                    for i in range(1, workbooks.Count + 1):
                        workbook = workbooks.Item(i)
                        if workbook.Name == name:
                            return workbook
            
            # 尝试app_match
            matched_object = self.app_match(object_name_list)
            print(f"DEBUG: app_match result = {matched_object}")
            
            if matched_object:
                for i in range(1, workbooks.Count + 1):
                    workbook = workbooks.Item(i)
                    if workbook.Name == matched_object:
                        print(f"DEBUG: Found matched workbook: {workbook.Name}")
                        return workbook
            
            # 如果都失败了，返回第一个工作簿
            if workbooks.Count > 0:
                first_workbook = workbooks.Item(1)
                print(f"DEBUG: Returning first workbook: {first_workbook.Name}")
                return first_workbook
            
            print("DEBUG: No workbooks available")
            return None
            
        except Exception as e:
            print(f"DEBUG: Exception in get_object_from_process_name: {e}")
            import traceback
            traceback.print_exc()
            return None

    def table2markdown(self, sheet_name: Union[str, int]) -> str:
        """
        Convert the table in the sheet to a markdown table string.
        :param sheet_name: The sheet name (str), or the sheet index (int), starting from 1.
        :return: The markdown table string.
        """

        sheet = self.com_object.Sheets(sheet_name)

        # Fetch the data from the sheet
        data = sheet.UsedRange()

        # Convert the data to a DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])

        # Drop the rows with all NaN values
        df = df.dropna(axis=0, how="all")

        # Convert the values to strings
        df = df.applymap(self.format_value)

        return df.to_markdown(index=False)

    def insert_excel_table(
        self, sheet_name: str, table: List[List[Any]], start_row: int, start_col: int
    ):
        """
        Insert a table into the sheet.
        :param sheet_name: The sheet name.
        :param table: The list of lists of values to be inserted.
        :param start_row: The start row.
        :param start_col: The start column.
        """
        sheet = self.com_object.Sheets(sheet_name)

        if str(start_col).isalpha():
            start_col = self.letters_to_number(start_col)

        for i, row in enumerate(table):
            for j, value in enumerate(row):
                sheet.Cells(start_row + i, start_col + j).Value = value

        return table

    def select_table_range(
        self,
        sheet_name: str,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
    ):
        """
        Select a range of cells in the sheet.
        :param sheet_name: The sheet name.
        :param start_row: The start row.
        :param start_col: The start column.
        :param end_row: The end row. If ==-1, select to the end of the document with content.
        :param end_col: The end column. If ==-1, select to the end of the document with content.
        """

        sheet_list = [sheet.Name for sheet in self.com_object.Sheets]
        if sheet_name not in sheet_list:
            print(
                f"Sheet {sheet_name} not found in the workbook, using the first sheet."
            )
            sheet_name = 1

        if str(start_col).isalpha():
            start_col = self.letters_to_number(start_col)

        if str(end_col).isalpha():
            end_col = self.letters_to_number(end_col)

        sheet = self.com_object.Sheets(sheet_name)

        if end_row == -1:
            end_row = sheet.Rows.Count
        if end_col == -1:
            end_col = sheet.Columns.Count

        try:
            sheet.Range(
                sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col)
            ).Select()
            return f"Range {start_row}:{start_col} to {end_row}:{end_col} is selected."
        except Exception as e:
            return f"Failed to select the range {start_row}:{start_col} to {end_row}:{end_col}. Error: {e}"

    def reorder_columns(self, sheet_name: str, desired_order: List[str] = None) -> str:
        """
        Reorder only non-empty columns based on desired_order.
        Empty columns remain in their original positions.
        :param sheet_name: Sheet to operate on
        :param desired_order: List of column header names to reorder
        :return: Success or error message
        """
        try:
            ws = self.com_object.Sheets(sheet_name)
            used_range = ws.UsedRange
            start_col = used_range.Column
            total_cols = used_range.Columns.Count
            last_col = start_col + total_cols - 1

            non_empty_columns = []
            empty_columns = []

            for col in range(1, last_col + 1):
                cell_value = ws.Cells(1, col).Value
                if cell_value and str(cell_value).strip():
                    non_empty_columns.append((str(cell_value).strip(), col))
                else:
                    empty_columns.append(col)

            print("📌 Non-empty columns:", [x[0] for x in non_empty_columns])
            print("📌 Empty columns at:", empty_columns)

            name_to_col = {name: col for name, col in non_empty_columns}

            column_data = []
            for name in desired_order:
                if name in name_to_col:
                    col_index = name_to_col[name]
                    data = []
                    row = 1
                    while True:
                        value = ws.Cells(row, col_index).Value
                        if value is None and row > 100:
                            break
                        data.append(value)
                        row += 1
                    column_data.append((name, data))
                else:
                    print(f"⚠️ Column '{name}' not found, skipping.")

            for _, col_index in sorted(non_empty_columns, key=lambda x: -x[1]):
                ws.Columns(col_index).Delete()

            insert_offset = 1
            for name, data in column_data:
                insert_pos = self.get_nth_non_empty_position(insert_offset, empty_columns)
                print(f"✅ Inserting '{name}' at position {insert_pos}")
                for row_index, value in enumerate(data, start=1):
                    ws.Cells(row_index, insert_pos).Value = value
                insert_offset += 1

            return f"Columns reordered successfully into: {desired_order}"

        except Exception as e:
            print(f"❌ Failed to reorder columns. Error: {e}")
            return f"Failed to reorder columns. Error: {e}"

    def get_range_values(
        self,
        sheet_name: str,
        start_row: int,
        start_col: int,
        end_row: int = -1,
        end_col: int = -1,
    ):
        """
        Get values from Excel sheet starting at (start_row, start_col) to (end_row, end_col).
        If end_row or end_col is -1, it automatically extends to the last used row or column.

        :param sheet_name: The name of the sheet.
        :param start_row: Starting row index (1-based)
        :param start_col: Starting column index (1-based)
        :param end_row: Ending row index, -1 means go to the last used row
        :param end_col: Ending column index, -1 means go to the last used column
        :return: List of lists (2D) containing the values
        """

        sheet_list = [sheet.Name for sheet in self.com_object.Sheets]
        if sheet_name not in sheet_list:
            print(
                f"Sheet {sheet_name} not found in the workbook, using the first sheet."
            )
            sheet_name = 1

        sheet = self.com_object.Sheets(sheet_name)

        used_range = sheet.UsedRange
        last_row = used_range.Row + used_range.Rows.Count - 1
        last_col = used_range.Column + used_range.Columns.Count - 1

        if end_row == -1:
            end_row = last_row
        if end_col == -1:
            end_col = last_col

        cell_range = sheet.Range(
            sheet.Cells(start_row, start_col), sheet.Cells(end_row, end_col)
        )

        try:
            values = cell_range.Value

        except Exception as e:
            return f"Failed to get values from range {start_row}:{start_col} to {end_row}:{end_col}. Error: {e}"

        # If it's a single cell, return [[value]]
        if not isinstance(values, tuple):
            return [[values]]

        # If it's a single row or column, make sure it’s 2D list
        if isinstance(values[0], (str, int, float, type(None))):
            return [list(values)]

        return [list(row) for row in values]

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

        excel_ext_to_fileformat = {
            ".xlsx": 51,  # Excel Workbook (default, no macros)
            ".xlsm": 52,  # Excel Macro-Enabled Workbook
            ".xlsb": 50,  # Excel Binary Workbook
            ".xls": 56,  # Excel 97-2003 Workbook
            ".xltx": 54,  # Excel Template
            ".xltm": 53,  # Excel Macro-Enabled Template
            ".csv": 6,  # CSV (comma delimited)
            ".txt": 42,  # Text (tab delimited)
            ".pdf": 57,  # PDF file (Excel 2007+)
            ".xps": 58,  # XPS file
            ".xml": 46,  # XML Spreadsheet 2003
            ".html": 44,  # HTML file
            ".htm": 44,  # HTML file
            ".prn": 36,  # Formatted text (space delimited)
        }

        # Enhanced path handling
        if not file_dir:
            file_dir = os.path.dirname(self.com_object.FullName)
        else:
            # Expand environment variables
            file_dir = os.path.expandvars(file_dir)
            # Expand user home directory
            file_dir = os.path.expanduser(file_dir)
            
        if not file_name:
            file_name = os.path.splitext(os.path.basename(self.com_object.FullName))[0]
            
        if not file_ext:
            file_ext = ".xlsx"  # Default to xlsx instead of csv

        # Validate and create directory if needed
        try:
            Path(file_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            # Fallback to user's Documents folder
            file_dir = os.path.join(os.path.expanduser("~"), "Documents")
            try:
                Path(file_dir).mkdir(parents=True, exist_ok=True)
            except Exception:
                # Final fallback to temp directory
                file_dir = tempfile.gettempdir()

        # Generate unique filename if file exists
        base_file_path = os.path.join(file_dir, file_name + file_ext)
        file_path = base_file_path
        counter = 1
        
        while os.path.exists(file_path):
            name_with_counter = f"{file_name}_{counter}"
            file_path = os.path.join(file_dir, name_with_counter + file_ext)
            counter += 1

        # Get file format
        file_format = excel_ext_to_fileformat.get(file_ext.lower(), 51)  # Default to xlsx

        try:
            # Attempt to save with different methods
            return self._attempt_save_with_fallback(file_path, file_format, file_ext)
            
        except Exception as e:
            return f"Failed to save the document to {file_path}. Error: {str(e)}"

    def _attempt_save_with_fallback(self, file_path: str, file_format: int, file_ext: str) -> str:
        """
        Attempt to save with multiple fallback strategies.
        """
        # Method 1: Standard SaveAs
        try:
            self.com_object.SaveAs(file_path, FileFormat=file_format)
            return f"Document successfully saved to {file_path}."
        except Exception as e1:
            print(f"Standard SaveAs failed: {e1}")
            
            # Method 2: SaveAs without FileFormat parameter (let Excel decide)
            try:
                self.com_object.SaveAs(file_path)
                return f"Document successfully saved to {file_path} (auto-format)."
            except Exception as e2:
                print(f"Auto-format SaveAs failed: {e2}")
                
                # Method 3: Export method for specific formats
                if file_ext.lower() in ['.pdf', '.xps']:
                    try:
                        export_format = 0 if file_ext.lower() == '.pdf' else 1  # 0=PDF, 1=XPS
                        self.com_object.ExportAsFixedFormat(
                            Type=export_format,
                            Filename=file_path,
                            Quality=0,  # 0=minimum, 1=maximum
                            IncludeDocProps=True,
                            IgnorePrintAreas=False,
                            OpenAfterPublish=False
                        )
                        return f"Document successfully exported to {file_path}."
                    except Exception as e3:
                        print(f"Export method failed: {e3}")
                
                # Method 4: Save to CSV and convert (for data formats)
                if file_ext.lower() in ['.csv', '.txt']:
                    try:
                        # Get current active sheet data
                        active_sheet = self.com_object.ActiveSheet
                        used_range = active_sheet.UsedRange
                        
                        if used_range:
                            # Save as CSV using pandas fallback
                            import pandas as pd
                            
                            # Get data from Excel
                            data = used_range.Value
                            if data:
                                if isinstance(data[0], (list, tuple)):
                                    df = pd.DataFrame(data[1:], columns=data[0])
                                else:
                                    df = pd.DataFrame([data])
                                
                                # Save using pandas
                                if file_ext.lower() == '.csv':
                                    df.to_csv(file_path, index=False)
                                else:
                                    df.to_csv(file_path, sep='\t', index=False)
                                    
                                return f"Document successfully saved to {file_path} (pandas method)."
                    except Exception as e4:
                        print(f"Pandas fallback failed: {e4}")
                
                # Final fallback: save as default Excel format in Documents
                try:
                    fallback_path = os.path.join(
                        os.path.expanduser("~"), "Documents", 
                        f"UFO_Excel_Export_{int(time.time())}.xlsx"
                    )
                    self.com_object.SaveAs(fallback_path, FileFormat=51)
                    return f"Document saved to fallback location: {fallback_path}"
                except Exception as e5:
                    return f"All save methods failed. Last error: {str(e5)}"

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
            ]
            
            for path in desktop_paths:
                if os.path.exists(path) and os.access(path, os.W_OK):
                    return path
            
            # Fallback to Documents
            return os.path.join(os.path.expanduser("~"), "Documents")
            
        except Exception:
            # Final fallback to temp directory
            return tempfile.gettempdir()


    @staticmethod
    def letters_to_number(letters: str) -> int:
        """
        Convert the column letters to the column number.
        :param letters: The column letters.
        :return: The column number.
        """
        number = 0
        for i, letter in enumerate(letters[::-1]):
            number += (ord(letter.upper()) - ord("A") + 1) * (26**i)
        return number

    @staticmethod
    def get_nth_non_empty_position(target_idx: int, empty_cols: List[int]) -> int:
        """
        Get the Nth available column index in the sheet, skipping empty columns.
        :param target_idx: The target index of the non-empty column.
        :param empty_cols: The list of empty column indexes.
        :return: The Nth non-empty column index.
        """
        col = 1
        non_empty_count = 0
        while True:
            if col not in empty_cols:
                non_empty_count += 1
            if non_empty_count == target_idx:
                return col
            col += 1

    @staticmethod
    def format_value(value: Any) -> str:
        """
        Convert the value to a formatted string.
        :param value: The value to be converted.
        :return: The converted string.
        """
        if isinstance(value, (int, float)):
            return "{:.0f}".format(value)
        return value

    @property
    def type_name(self):
        return "COM/EXCEL"

    @property
    def xml_format_code(self) -> int:
        return 46


@ExcelWinCOMReceiver.register
class GetSheetContentCommand(WinCOMCommand):
    """
    The command to insert a table.
    """

    def execute(self):
        """
        Execute the command to insert a table.
        :return: The inserted table.
        """
        return self.receiver.table2markdown(self.params.get("sheet_name"))

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "table2markdown"


@ExcelWinCOMReceiver.register
class InsertExcelTableCommand(WinCOMCommand):
    """
    The command to insert a table.
    """

    def execute(self):
        """
        Execute the command to insert a table.
        :return: The inserted table.
        """
        return self.receiver.insert_excel_table(
            sheet_name=self.params.get("sheet_name", 1),
            table=self.params.get("table"),
            start_row=self.params.get("start_row", 1),
            start_col=self.params.get("start_col", 1),
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "insert_excel_table"


@ExcelWinCOMReceiver.register
class SelectTableRangeCommand(WinCOMCommand):
    """
    The command to select a table.
    """

    def execute(self):
        """
        Execute the command to select a table.
        :return: The selected table.
        """
        return self.receiver.select_table_range(
            sheet_name=self.params.get("sheet_name", 1),
            start_row=self.params.get("start_row", 1),
            start_col=self.params.get("start_col", 1),
            end_row=self.params.get("end_row", 1),
            end_col=self.params.get("end_col", 1),
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "select_table_range"


@ExcelWinCOMReceiver.register
class GetRangeValuesCommand(WinCOMCommand):
    """
    The command to get values from a range.
    """

    def execute(self):
        """
        Execute the command to get values from a range.
        :return: The values from the range.
        """
        return self.receiver.get_range_values(
            sheet_name=self.params.get("sheet_name", 1),
            start_row=self.params.get("start_row", 1),
            start_col=self.params.get("start_col", 1),
            end_row=self.params.get("end_row", -1),
            end_col=self.params.get("end_col", -1),
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "get_range_values"


@ExcelWinCOMReceiver.register
class ReorderColumnsCommand(WinCOMCommand):
    """
    The command to reorder columns in a sheet.
    """

    def execute(self):
        """
        Execute the command to reorder columns in a sheet.
        :return: The result of reordering columns.
        """
        return self.receiver.reorder_columns(
            sheet_name=self.params.get("sheet_name", 1),
            desired_order=self.params.get("desired_order"),
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "reorder_columns"


@ExcelWinCOMReceiver.register
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
            return "Error: No active Excel workbook found. Please ensure Excel is running with an open workbook."
        
        # Get parameters with enhanced defaults
        file_dir = self.params.get("file_dir", "")
        file_name = self.params.get("file_name", "")
        file_ext = self.params.get("file_ext", ".xlsx")
        
        # Handle special path cases
        if file_dir and "%USERNAME%" in file_dir:
            # Replace %USERNAME% with actual username
            import getpass
            username = getpass.getuser()
            file_dir = file_dir.replace("%USERNAME%", username)
        
        # Ensure file extension starts with dot
        if file_ext and not file_ext.startswith('.'):
            file_ext = '.' + file_ext
            
        # Use safe desktop path if specified but invalid
        if file_dir and "Desktop" in file_dir and not os.path.exists(file_dir):
            file_dir = self.receiver.get_safe_desktop_path()
        
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
