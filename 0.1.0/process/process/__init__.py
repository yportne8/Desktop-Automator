import win32gui
import traceback
from typing import Union
from pathlib import Path
from abc import ABC, abstractmethod
from win32ui import CreateFileDialog

import pandas as pd

from .office import FileManager, Excel, Word, Outlook, SharePoint, Access
from .shared import Table, Console, Browser, Notify, WindowManager
from .scheduler import Task, Scheduler


class Process(ABC):


    __all__ = ["dataIn", "dataOut", "main", "get_excel", "get_word",
               "get_outlook", "get_access", "start_task", "schedule_task"
               "download_sharepoint_folder", "download_sharepoint_file",
               "sharepoint_folder_contents", "sharepoint_file_to_dataframe", 
               "local_file_to_dataframe", "export_dataframe_to_worksheet",
               "execute_workbook_macro"]


    def __init__(self, username: str=None, password: str=None):
        """ 
        _[Capabilities of the Process Class]
         self.sharepoint contains functions to stream Excel files
         directly into a pd.DataFrame so that edits can be made 
         using the static editing functions attacheded to Table.
        
         self.excel houses additional functions to import and export
         local files in various format (csv, xlsx, pkl) tp/from a 
         DataFrame. self.excel can also load the DataFrame into a
         Worksheet, then execute a macro housed in the Worbook.

         self.filemanager can resolve local paths nested inside OneDrive,
         and even if the requested download path already exists. The 
         filemanager holds, as properties, often used file locations: 
         Desktop, Documents, and OneDrive folders. The filemanager can
         also create secured temporary files, as well as temp files in
         memory.
        
         Outlook contains prebuilt one-line functions to draft and send
         emails.
        
         Word can open documents (with or without encryption). After 
         opening a Word document via get_desktop_word, the returned 
         document object can be operated on.
         
         ```python
         word = self.get_desktop_word()
         document = word.open("C:\\Path\to\document.docx")
         document.Write("Hello World!")
         ````
        
         Two monitors (monitoring sharepoint file/folder changes and
         monitoring the Outlook inbox) and a timed scheulder is avaiable
         to create triggers to automate the execution of the script(s).
         
         self.window.title can be reassigned with any partial window
         title to take control over the running application's window, 
         including: hiding, unhiding, moving, resizing, keeping-on-top, 
         and sending virtual keyboard strokes directly onto the window
         to automate execution of scripts with attached GUIs.
        
         User notifications are built in using self.notify, with
         the option of notification via win native messagebox
         or a taskbar popup._
        """

        self.editor = Table
        self.sharepoint = SharePoint(username, password)         
        self.filemanager = FileManager()
        self.window = WindowManager()
        self.scheduler = Scheduler()
        self.webbrowser = Browser() 
        self.outlook = Outlook()
        self.console = Console()
        self.notify = Notify()
        self.excel = Excel()
        self.word = Word()
        
    
    @abstractmethod
    def dataIn(self, *args, **kwargs):
        pass

    
    @abstractmethod
    def dataOut(self, *args, **kwargs):
        pass

    def get_desktop_excel(self):
        return self.excel.app

    def get_desktop_word(self):
        return self.word.app

    def get_desktop_outlook(self):
        return self.word.app

    def get_desktop_access():
        return Access() # to be completed...

    def start_task(self, specifications: dict):
        """_Creates and starts a scheduler.Task_

        Args:
            specifications (dict): _task specifications_
        """

        try:
            task = Task(specifications)
            task.main()
        
        except:
            Task()

    def schedule_task(self, task: Task):
        """_Schedules a scheduler.Task._

        Args:
            task (_Task_): _scheduler.Task_
        """

        try:
            self.scheduler.add_task(task)
        
        except:
            Scheduler()

    def download_sharepoint_folder(self, url_or_list_of_folders_from_shared_documents: Union[str, list],
                                   destination: Union[str, Path]=None) -> str:
        """_Download a sharepoint folder. Default download destination is the user's Downloads directory._

        Args:
            url_or_list_of_folders_from_shared_documents (Union[str, list]): _as named_
            destination (Union[str, Path], optional): _local parent folder_. Defaults to None.

        Returns:
            str: _description_
        """

        try:
            return str(self.sharepoint.doc_folder_download(url_or_list_of_folders_from_shared_documents,
                                                           destination))
        
        except Exception as e:
            traceback.print_exception(e)
            print("Has self.site been changed to the corrected SharePoint site?")

    def download_sharepoint_file(self, url_or_list_of_folders_from_shared_documents: Union[str, list],
                                 destination: Union[str, Path]=None) -> str:
        """_Download a sharepoint file. Default download destination is the user's Downloads directory._

        Args:
            url_or_list_of_folders_from_shared_documents (Union[str, list]): _as named_
            dest (Union[str, Path], optional): _local parent folder_. Defaults to None.

        Returns:
            str: _description_
        """

        try:
            return str(self.sharepoint.doc_folder_download(url_or_list_of_folders_from_shared_documents,
                                                           destination))
        
        except Exception as e:
            traceback.print_exception(e)
            print("Has self.site been changed to the corrected SharePoint site?")

    def sharepoint_folder_content_names(self, url_or_list_of_folders_from_shared_documents: Union[str, list]) -> dict:
        """_Returns a dictionary of folder and file name. The names included in either list (folders, files),
            can be added to the [list, of, directories, to, asset], if passed to the url... parameter,
            to download the desired asset:
            
            ```python
            downloadPath = self.sharepoint.doc_file_download([list, of, directories, to, asset, file.xlsx])
            ```_

        Args:
            url_or_list_of_folders_from_shared_documents (Union[str, list]): _as named_

        Returns:
            dict: _folders and files_
        """

        try:
            xfolders = self.sharepoint.doc_folders(url_or_list_of_folders_from_shared_documents)
            xfiles = self.sharepoint.doc_files(url_or_list_of_folders_from_shared_documents)
  
            return {"folders": [f.name for f in xfolders],
                    "files"  : [f.name for f in xfiles]}
  
        except Exception as e:
            traceback.print_exception(e)
            print("Has self.site been changed to the corrected SharePoint site?")

    def sharepoint_folder_content_urls(self, url_or_list_of_folders_from_shared_documents: Union[str, list]) -> dict:
        # [TODO] a nested dictionary of asset names and urls
        pass

    def request_file_location(self, prompt: str) -> str:
        """_Request file location from the user._

        Args:
            prompt (str): _user prompt_

        Returns:
            str: _c:\\path\\to\\file.xlsx_
        """

        print(prompt)
        
        fileDialog = CreateFileDialog(False,None,None,False)
        fileDialog.DoModal() # blocking, required
        
        return fileDialog.GetPathName()

    def request_directory_location(self, prompt: str, starting_folder: str="documents") -> str:
        """_Request directory location from the user._

        Args:
            prompt (str): _user prompt_
            starting_folder (str, optional): _desktop or documents_. Defaults to "documents".

        Returns:
            _str_: _c:\\path\\to\\asset_
        """

        try:
            pidl = {
                "desktop": win32gui.shell.SHGetFolderLocation (0,win32gui.shell.shellcon.CSIDL_DESKTOP,0,0),
                "documents": win32gui.shell.SHGetFolderLocation (0,win32gui.shell.shellcon.CSIDL_PERSONAL,0,0)}[starting_folder]
        
        except:
            print("starting_folder options: desktop or documents")
        
        print(prompt)
        pidl,_,_ = win32gui.shell.SHBrowseForFolder (win32gui.GetDesktopWindow(),pidl,prompt,0,None,None)
        return win32gui.shell.SHGetPathFromIDList(pidl)
    
    def execute_excel_macro(self, workbook: Union[Path, str], module: str, macro: str):
        """_Executes a workbook macro. Please note that this process is blocking until completed._

        Args:
            workbook (Union[Path, str]): _description_
            module (str): _description_
            macro (str): _description_
        """

        workbook = Path(workbook)
        
        if workbook.exists():
            xl = self.get_excel()
            
            try:
                xl.exec_macro(workbook, module, macro)
            
            except Exception as e:
                traceback.print_exception(e)

            # xl is closed when garbage collected after function call.
        
        else:
            print(f"Could not find: {str(workbook)}.")


    def main(self):
        """
        _Connects dataIn and dataOut into single uninterrupted process._
        """

        parameters = self.dataIn()
        if parameters:

            if type(parameters) == Exception:
                traceback.print_exception(parameters)

            elif type(parameters) == list:
                try:
                    e = self.dataOut(*parameters)
                    if e:
                        traceback.print_exception(e)
                    
                    else:
                        print("DataIn Complete!")
            
                except Exception as e:
                    print("Exception raised for dataOut",e)
            
            elif type(parameters) in type(dict):
                e = self.dataOut(**parameters)
                if e:
                    traceback.print_exception(e)
                    
                else:
                    print("Process Complete!")
        
            else:
                print(f"Exec parameters must be type list or dict, not type: {type(parameters)}")
                
        else:
            e = self.dataOut(*parameters)
            if e:
                traceback.print_exception(e)
            
            else:
                print("Process Complete!")