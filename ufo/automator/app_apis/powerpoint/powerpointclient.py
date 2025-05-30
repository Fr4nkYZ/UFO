# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import os
import time
import tempfile
import getpass
from pathlib import Path
from typing import Dict, Type, List

from ufo.automator.app_apis.basic import WinCOMCommand, WinCOMReceiverBasic
from ufo.automator.basic import CommandBasic


class PowerPointWinCOMReceiver(WinCOMReceiverBasic):
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
            print(f"DEBUG: Starting PowerPoint get_object_from_process_name")
            
            if not self.client:
                print("DEBUG: PowerPoint client is None, returning None")
                return None
            
            # 获取所有演示文稿
            presentations = self.client.Presentations
            print(f"DEBUG: Presentations collection count = {presentations.Count}")
            
            object_name_list = []
            for i in range(1, presentations.Count + 1):
                try:
                    presentation = presentations.Item(i)
                    name = presentation.Name
                    object_name_list.append(name)
                    print(f"DEBUG: Presentation {i}: {name}")
                    print(f"DEBUG: Presentation {i} encoding: {name.encode('utf-8')}")
                except Exception as e:
                    print(f"DEBUG: Error getting presentation {i}: {e}")
            
            print(f"DEBUG: object_name_list = {object_name_list}")
            print(f"DEBUG: process_name = {self.process_name}")
            print(f"DEBUG: process_name encoding: {self.process_name.encode('utf-8')}")
            
            # 如果没有打开的演示文稿，创建一个新演示文稿
            if presentations.Count == 0:
                print("DEBUG: No presentations open, creating a new presentation")
                try:
                    new_presentation = self.client.Presentations.Add()
                    print(f"DEBUG: New presentation created: {new_presentation.Name}")
                    return new_presentation
                except Exception as e:
                    print(f"DEBUG: Failed to create new presentation: {e}")
                    return None
            
            # 如果有演示文稿，尝试匹配
            matched_object = self.app_match(object_name_list)
            print(f"DEBUG: app_match result = {matched_object}")
            
            # 如果process_name只是"PowerPoint"或匹配失败，返回第一个演示文稿
            if (self.process_name.lower() in ["powerpoint", "pptx", "microsoft powerpoint"] or 
                not matched_object):
                if presentations.Count > 0:
                    first_presentation = presentations.Item(1)
                    print(f"DEBUG: Returning first presentation: {first_presentation.Name}")
                    return first_presentation
            
            # 尝试找到匹配的演示文稿
            for presentation in presentations:
                if presentation.Name == matched_object:
                    print(f"DEBUG: Found matched presentation: {presentation.Name}")
                    return presentation
            
            # 如果都失败了，返回第一个演示文稿
            if presentations.Count > 0:
                first_presentation = presentations.Item(1)
                print(f"DEBUG: Fallback: returning first presentation: {first_presentation.Name}")
                return first_presentation
            
            print("DEBUG: No presentations available and failed to create new one")
            return None
            
        except Exception as e:
            print(f"DEBUG: Exception in PowerPoint get_object_from_process_name: {e}")
            import traceback
            traceback.print_exc()
            return None

    def set_background_color(self, color: str, slide_index: List[int] = None) -> str:
        """
        Set the background color of the slide(s).
        :param color: The hex color code (in RGB format) to set the background color.
        :param slide_index: The list of slide indexes to set the background color. If None, set the background color for all slides.
        :return: The result of setting the background color.
        """
        try:
            if self.com_object is None:
                return "Error: No active PowerPoint presentation found."
            
            if not slide_index:
                slide_index = range(1, self.com_object.Slides.Count + 1)

            # Remove '#' if present and validate hex color
            if color.startswith('#'):
                color = color[1:]
            
            if len(color) != 6:
                return f"Invalid hex color code: {color}. Please use 6-digit hex format (e.g., 'FF0000' for red)."

            red = int(color[0:2], 16)
            green = int(color[2:4], 16)
            blue = int(color[4:6], 16)
            bgr_hex = (blue << 16) + (green << 8) + red

            modified_slides = []
            for index in slide_index:
                if index < 1 or index > self.com_object.Slides.Count:
                    print(f"DEBUG: Skipping invalid slide index: {index}")
                    continue
                
                try:
                    slide = self.com_object.Slides(index)
                    slide.FollowMasterBackground = False
                    slide.Background.Fill.Visible = True
                    slide.Background.Fill.Solid()
                    slide.Background.Fill.ForeColor.RGB = bgr_hex  # PowerPoint uses BGR format
                    modified_slides.append(index)
                except Exception as e:
                    print(f"DEBUG: Failed to set background for slide {index}: {e}")
                    
            if modified_slides:
                return f"Successfully set the background color to #{color} for slide(s) {modified_slides}."
            else:
                return f"Failed to set background color for any slides."
                
        except Exception as e:
            return f"Failed to set the background color. Error: {e}"

    def save_as(
        self, file_dir: str = "", file_name: str = "", file_ext: str = "", current_slide_only: bool = False
    ) -> str:
        """
        Save the document to other formats with enhanced error handling.
        :param file_dir: The directory to save the file.
        :param file_name: The name of the file without extension.
        :param file_ext: The extension of the file.
        :param current_slide_only: Whether to save only the current slide (for image formats).
        :return: Success message or error details.
        """

        ppt_ext_to_fileformat = {
            ".pptx": 24,  # PowerPoint Presentation (OpenXML)
            ".ppt": 0,  # PowerPoint 97-2003 Presentation
            ".pdf": 32,  # PDF file
            ".xps": 33,  # XPS file
            ".potx": 25,  # PowerPoint Template (OpenXML)
            ".pot": 5,  # PowerPoint 97-2003 Template
            ".ppsx": 27,  # PowerPoint Show (OpenXML)
            ".pps": 1,  # PowerPoint 97-2003 Show
            ".odp": 35,  # OpenDocument Presentation
            ".jpg": 17,  # JPG images (slides exported as .jpg)
            ".png": 18,  # PNG images
            ".gif": 19,  # GIF images
            ".bmp": 20,  # BMP images
            ".tif": 21,  # TIFF images
            ".tiff": 21,  # TIFF images
            ".rtf": 6,  # Outline RTF
            ".html": 12,  # Single File Web Page
            ".mp4": 39,  # MPEG-4 video (requires PowerPoint 2013+)
            ".wmv": 38,  # Windows Media Video
            ".xml": 10,  # PowerPoint 2003 XML Presentation
        }

        ppt_ext_to_formatstr = {
            ".jpg": "JPG",  # JPG images (slides exported as .jpg)
            ".png": "PNG",  # PNG images
            ".gif": "GIF",  # GIF images
            ".bmp": "BMP",  # BMP images
            ".tif": "TIF",  # TIFF images
            ".tiff": "TIF",  # TIFF images
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
                file_name = f"powerpoint_presentation_{int(time.time())}"
            
        if not file_ext:
            file_ext = ".pptx"  # Default to pptx
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

        try:
            # Check if com_object exists
            if self.com_object is None:
                return "Error: No active PowerPoint presentation found."
            
            # Attempt to save with different methods
            return self._attempt_save_with_fallback(file_path, file_ext, current_slide_only, ppt_ext_to_fileformat, ppt_ext_to_formatstr)
            
        except Exception as e:
            return f"Failed to save the presentation to {file_path}. Error: {str(e)}"

    def _attempt_save_with_fallback(self, file_path: str, file_ext: str, current_slide_only: bool, format_map: dict, export_map: dict) -> str:
        """
        Attempt to save with multiple fallback strategies.
        """
        # Method 1: Export single slide (for image formats)
        if file_ext in export_map.keys():
            try:
                if self.com_object.Slides.Count == 1 or current_slide_only:
                    if current_slide_only:
                        try:
                            # Try to get current slide from slideshow
                            current_slide_idx = self.com_object.SlideShowWindow.View.Slide.SlideIndex
                        except:
                            # If no slideshow is running, use first slide
                            current_slide_idx = 1
                    else:
                        current_slide_idx = 1
                    
                    self.com_object.Slides(current_slide_idx).Export(
                        file_path, export_map.get(file_ext, "PNG")
                    )
                    
                    if os.path.exists(file_path):
                        file_size = os.path.getsize(file_path)
                        return f"Slide {current_slide_idx} successfully exported to {file_path} (Size: {file_size} bytes)"
                    else:
                        return f"Export completed but file not found at {file_path}"
                        
                else:
                    # Export all slides as individual files
                    base_name = os.path.splitext(file_path)[0]
                    exported_files = []
                    
                    for i in range(1, self.com_object.Slides.Count + 1):
                        slide_path = f"{base_name}_slide_{i}{file_ext}"
                        self.com_object.Slides(i).Export(slide_path, export_map.get(file_ext, "PNG"))
                        if os.path.exists(slide_path):
                            exported_files.append(slide_path)
                    
                    if exported_files:
                        return f"Successfully exported {len(exported_files)} slides to {os.path.dirname(file_path)}"
                    else:
                        return "Export completed but no files found"
                        
            except Exception as e1:
                print(f"DEBUG: Export method failed: {e1}")
        
        # Method 2: Standard SaveAs
        try:
            file_format = format_map.get(file_ext.lower(), 24)  # Default to pptx
            self.com_object.SaveAs(file_path, FileFormat=file_format)
            
            if os.path.exists(file_path):
                file_size = os.path.getsize(file_path)
                return f"Presentation successfully saved to {file_path} (Size: {file_size} bytes)"
            else:
                return f"SaveAs completed but file not found at {file_path}"
                
        except Exception as e2:
            print(f"DEBUG: Standard SaveAs failed: {e2}")
            
            # Method 3: SaveAs without FileFormat parameter (let PowerPoint decide)
            try:
                self.com_object.SaveAs(file_path)
                if os.path.exists(file_path):
                    return f"Presentation successfully saved to {file_path} (auto-format)"
                else:
                    return f"Auto-format save completed but file not found"
            except Exception as e3:
                print(f"DEBUG: Auto-format SaveAs failed: {e3}")
                
                # Method 4: Export as fixed format for PDF/XPS
                if file_ext.lower() in ['.pdf', '.xps']:
                    try:
                        export_format = 2 if file_ext.lower() == '.pdf' else 4  # ppFixedFormatTypePDF, ppFixedFormatTypeXPS
                        self.com_object.ExportAsFixedFormat(
                            Path=file_path,
                            FixedFormatType=export_format,
                            Intent=1,  # ppFixedFormatIntentPrint
                            FrameSlides=False,
                            HandoutOrder=1,  # ppPrintHandoutOrderHorizontalFirst
                            OutputType=1,  # ppPrintOutputSlides
                            PrintHiddenSlides=False,
                            PrintRange=None,
                            RangeType=1,  # ppPrintAll
                            SlideShowName="",
                            IncludeDocProps=True,
                            KeepIRMSettings=True,
                            DocStructureTags=True,
                            BitmapMissingFonts=True,
                            UseISO19005_1=False
                        )
                        
                        if os.path.exists(file_path):
                            return f"Presentation successfully exported to {file_path}"
                        else:
                            return f"Export completed but file not found"
                    except Exception as e4:
                        print(f"DEBUG: Export as fixed format failed: {e4}")
                
                # Method 5: Copy and save approach
                try:
                    # Create a new presentation and copy content
                    new_presentation = self.client.Presentations.Add()
                    
                    # Copy all slides from original presentation
                    for i in range(1, self.com_object.Slides.Count + 1):
                        self.com_object.Slides(i).Copy()
                        new_presentation.Slides.Paste()
                    
                    # Remove the default empty slide
                    if new_presentation.Slides.Count > self.com_object.Slides.Count:
                        new_presentation.Slides(1).Delete()
                    
                    # Save the new presentation
                    new_presentation.SaveAs(file_path, FileFormat=format_map.get(file_ext.lower(), 24))
                    new_presentation.Close()
                    
                    if os.path.exists(file_path):
                        return f"Presentation successfully saved to {file_path} (copy method)"
                    else:
                        return f"Copy method completed but file not found"
                        
                except Exception as e5:
                    print(f"DEBUG: Copy method failed: {e5}")
                
                # Final fallback: save as default PowerPoint format in Documents
                try:
                    fallback_path = os.path.join(
                        os.path.expanduser("~"), "Documents", 
                        f"UFO_PowerPoint_Export_{int(time.time())}.pptx"
                    )
                    self.com_object.SaveAs(fallback_path, FileFormat=24)
                    return f"Presentation saved to fallback location: {fallback_path}"
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
        return "COM/POWERPOINT"

    @property
    def xml_format_code(self) -> int:
        return 10


@PowerPointWinCOMReceiver.register
class SetBackgroundColorCommand(WinCOMCommand):
    """
    The command to set the background color of the slide(s).
    """

    def execute(self):
        """
        Execute the command to set the background color of the slide(s).
        :return: The result of setting the background color.
        """
        if self.receiver.com_object is None:
            return "Error: No active PowerPoint presentation found. Please ensure PowerPoint is running with an open presentation."
        
        return self.receiver.set_background_color(
            self.params.get("color", ""), 
            self.params.get("slide_index", [])
        )

    @classmethod
    def name(cls) -> str:
        """
        The name of the command.
        """
        return "set_background_color"


@PowerPointWinCOMReceiver.register
class SaveAsCommand(WinCOMCommand):
    """
    The command to save the document to various formats.
    """

    def execute(self):
        """
        Execute the command to save the document to specified format.
        :return: The result of saving the document.
        """
        # 首先检查com_object是否为None
        if self.receiver.com_object is None:
            return "Error: No active PowerPoint presentation found. Please ensure PowerPoint is running with an open presentation."
        
        # Get parameters with enhanced defaults
        file_dir = self.params.get("file_dir", "")
        file_name = self.params.get("file_name", "")
        file_ext = self.params.get("file_ext", ".pptx")
        current_slide_only = self.params.get("current_slide_only", False)
        
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
                current_slide_only=current_slide_only,
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
