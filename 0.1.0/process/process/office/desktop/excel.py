import traceback
import pythoncom
from io import BytesIO
from typing import Union
from pathlib import Path
from win32com import client
from threading import Thread

import msoffcrypto
import pandas as pd


class FailedToDecrypt(Exception):
    pass


class FailedToOpen(Exception):
    pass


class Excel:

    """
    _A light wrapper around Excel that integrates pandas.DataFrames._
    """

    def __init__(self):
        self._df = pd.DataFrame()
        self.thread = None
    
    def __del__(self):
        if self.thread:
            try:
                self.thread.join()
            except:
                pass

        try:
            # blocking until closed
            self.app.Application.Quit()
            print("Excel has closed.")
        except Exception as e:
            traceback.print_exception(e)

    @property
    def df(self):
        # [NOTE] DataFrames, even dataframe splices
        # should be copied for table operations
        # then reassigned after operation.
        return self._df.copy()

    @df.setter
    def df(self, value: pd.DataFrame):
        self._df = value

    @property
    def app(self):
        return client.Dispatch("Excel.Application")
    
    def _is_alive(self):
        if self.thread:
            try:
                return self.thread.is_alive()
            except Exception as e:
                print("Threading Exception",e)
                print("Attempting to shutdown Excel...")
                try:
                    self.thread.join()
                except:
                    pass
                self.thread = None
                self.app.Application.Quit()
        else:
            return False

    def _get_local__as_dataframe(self, path: Union[str, Path] = None,
                                    sheetNameOrNum: Union[str, int] = 0,
                                    password: str = None) -> pd.DataFrame:
        """_Returns a data frame from the requested source._

        Args:
            path (Union[str, Path]): _description_
            sheetNameOrNum (Union[str, int], optional): _description_. Defaults to 0.
            password (str, optional): _description_. Defaults to None.

        Raises:
            FileNotFoundError: _as named_

        Returns:
            pandas.DataFrame: _pd.DataFrame_
        """

        path = Path(path)
        if not path.exists():
            raise FileNotFoundError
        
        try:
            if not path: return pd.read_clipboard()
            
            if path.suffix == ".csv": return pd.read_csv(path)
            
            if path.suffix == ".pkl": return pd.read_pickle(path)
            
            if path.suffix in [".xlsx", ".xls", ".xlsm"]:
            
                if password:
                    iobuffer = BytesIO()
                    with open(path, "rb") as f:
                        contents = msoffcrypto.OfficeFile(f)
                        contents.load_key(password=password)
                        contents.decrypt(iobuffer)
                    return pd.read_excel(iobuffer)
            
                else:
                    return pd.read_excel(path, sheet_name=sheetNameOrNum)
            
            else:
                print("Unknown file type.")
        
        except Exception as e:
            traceback.print_exception(e)
            print("Failed to read file.")

    def _load_dataframe(self, path: Union[str,Path], df: pd.DataFrame,
                        sheet: Union[str, int]=None, delete_if_existing: bool=False):
        """_Loads a pandas dataframe into an Excel Worksheet._

        Args:
            path (Union[str,Path]): _path to workbook_
            sheet_name (str): _worksheet name_
            df (pd.DataFrame, optional): _dataframe, uses self.df store if not passed_. Defaults to None.
            delete_if_existing (bool, optional): _overwrite the worksheet if existing_. Defaults to False.
        """

        path = Path(path)
        
        if path.suffix == ".csv":
            if path.exists():
                if delete_if_existing:
                    df.to_csv(path)
                else:
                    print("Nothing done. File already exists.")
            else:
                df.to_csv(path)
            return

        if path.exists() and not sheet:
            print("A sheet name or index is required to add a DataFrame to a Workbook.")
            return

        if not path.exists():
            workbook = self.app.Workbooks.Add()
            sheet = 1 # [NOTE] sheet index starts at 1

        workbook = self.app.Workbooks.Open(str(path))
        shtNames = [wksht.Name for wksht in workbook.Worksheets]
        shtNameExists = False
        if type(sheet) == int:
            # [NOTE] + 2: sheet index starts at 1, then after the last.
            if sheet != len(shtNames) + 2: shtNameExists = True
        else:
            if sheet in shtNames: shtNameExists = True
        
        if (shtNameExists and delete_if_existing) or not shtNameExists: 
                df.to_excel(self.app, sheet_name=sheet)
        else:
            print("Nothing done. Worksheet already exists.")
        
        workbook.Save()
        workbook.Close()
    
    def _open_encrypted(self, strpath, undatelinks, readonly, password):
        """_Directs Excel to open a encrypted Workbook and return the opened
            Workbook obj._

        Args:
            strpath (_type_): _description_
            undatelinks (_type_): _description_
            readonly (_type_): _description_
            password (_type_): _description_

        Raises:
            FailedToDecrypt: _description_

        Returns:
            _type_: _description_
        """

        try:
            file_format_declaration=None
            return self.app.Workbooks.Open(strpath,undatelinks,readonly,file_format_declaration,password)
        
        except Exception as e:
            traceback.print_exception(e)
            raise FailedToDecrypt

    def open(self, path: Union[str, Path], updatelinks: bool = False,
             readonly: bool = False, password: str = None):
        """_Open a workbook, operations can be performed on the workbook object after open._

        Args:
            path (Union[str, Path]): _C:\\path\to\workbook.xlsx_
            updatelinks (bool, optional): _as named_. Defaults to False.
            readonly (bool, optional): _as named_. Defaults to False.
            password (str, optional): _as named_. Defaults to None.

        Raises:
            FileNotFoundError: _as named_
        """

        if password:
            # auto open encrypted
            return self._open_encrypted(str(path), readonly, password)
        
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError
        
        # [TODO] The difference between editable and readonly.
        try:
            return self.app.Workbooks.Open(str(path), updatelinks, readonly)
        
        except Exception as e:
            traceback.print_exception(e)
    
    def hide(self):
        """
        _Hides the window, does not close it. New windows cannot be 
         opened without unhiding all other windows._
        """
        
        self.app.Visible = False 
    
    def unhide(self):
        """
        _Unhide the window._
        """
        
        self.app.Visible = True

    def copy_table_to_clipboard_as_image(self, workbookName, worksheetName, range_start: str, range_end: str):
        """_Copies a table from the active workbook/worksheet and makes it available through
            win32com, the table is not saved to clipboard and nothing is returned from this function._

        Args:
            range_start (str): _table range start_
            range_end (str): _table range end_
        """

        # [TODO] set properties like word, including font changes and highlights
        # add a function to detect highlights in a data range.
        # save new workbooks to user TEMP???? this might be a limitation...
        try:
            workbook = self.app.Workbooks[workbookName]
        except:
            print(f"{workbookName} not found.")

        try:
            worksheet = workbook.Worksheets[worksheetName]
        except:
            print(f"{worksheetName} not found.")

        range = f"{range_start.upper()}:{range_end.upper()}"
        worksheet.Range(range).CopyPicture()

    def get_embedded_html_table_as_dataframe(self, url: str) -> pd.DataFrame:
        """_A wrapper around pd.read_html to catch Exception thrown
            by webpage with no readable table._

        Args:
            url (str): _webpage url_

        Returns:
            _pandas.DataFrame_: _pd.DataFrame_
        """
        
        try:
            return pd.read_html(url)[0]
        
        except:
            print("Could not find a readable table.")

    def get_local_xlsx_as_dataframe(self, path: Union[str, Path] = None,
                                    sheetNameOrNum: Union[str, int] = 0,
                                    password: str = None) -> pd.DataFrame:
        """_Returns a data frame from the requested source._

        Args:
            path (Union[str, Path]): _description_
            sheetNameOrNum (Union[str, int], optional): _description_. Defaults to 0.
            password (str, optional): _description_. Defaults to None.

        Raises:
            FileNotFoundError: _as named_

        Returns:
            pandas.DataFrame: _pd.DataFrame_
        """
        
        return self._get_local__as_dataframe(path, sheetNameOrNum, password)

    def get_local_csv_as_dataframe(self, path: Union[str, Path] = None,
                                    sheetNameOrNum: Union[str, int] = 0,
                                    password: str = None) -> pd.DataFrame:
        """_Returns a data frame from the requested source._

        Args:
            path (Union[str, Path]): _description_
            sheetNameOrNum (Union[str, int], optional): _description_. Defaults to 0.
            password (str, optional): _description_. Defaults to None.

        Raises:
            FileNotFoundError: _as named_

        Returns:
            pandas.DataFrame: _pd.DataFrame_
        """
        
        return self._get_local__as_dataframe(path, sheetNameOrNum, password)
    
    def get_local_pkl_as_dataframe(self, path: Union[str, Path] = None,
                                    sheetNameOrNum: Union[str, int] = 0,
                                    password: str = None) -> pd.DataFrame:
        """_Returns a data frame from the requested source._

        Args:
            path (Union[str, Path]): _description_
            sheetNameOrNum (Union[str, int], optional): _description_. Defaults to 0.
            password (str, optional): _description_. Defaults to None.

        Raises:
            FileNotFoundError: _as named_

        Returns:
            pandas.DataFrame: _pd.DataFrame_
        """
        
        return self._get_local__as_dataframe(path, sheetNameOrNum, password)

    def load_dataframe_into_worksheet(self, path: Union[str,Path], sheet_name: str,
                                      df: pd.DataFrame=None, delete_if_existing: bool=False):
        """_Loads a pandas dataframe into an Excel Worksheet._

        Args:
            path (Union[str,Path]): _path to workbook_
            sheet_name (str): _worksheet name_
            df (pd.DataFrame, optional): _dataframe, uses self.df store if not passed_. Defaults to None.
            delete_if_existing (bool, optional): _overwrite the worksheet if existing_. Defaults to False.
        """

        self._load_dataframe(path, sheet_name, df, delete_if_existing)

    def load_dataframe_into_csv(self, path: Union[str,Path], df: pd.DataFrame=None,
                                delete_if_existing: bool=False):
        """_Loads a pandas dataframe into an Excel Worksheet._

        Args:
            path (Union[str,Path]): _path to workbook_
            sheet_name (str): _worksheet name_
            df (pd.DataFrame, optional): _dataframe, uses self.df store if not passed_. Defaults to None.
            delete_if_existing (bool, optional): _overwrite the worksheet if existing_. Defaults to False.
        """

        __sheet_name = 1
        self._load_dataframe(path, __sheet_name, df, delete_if_existing)

    def exec_macro(self, path: Union[str,Path], module:str, macro: str,
                  readonly: bool = True, password: str = None):
        """_Run is blocking until macro completion. If excel needs to be 
            hidden during this process, this needs to happen from within
            the macro._

        Args:
            path (Path): _C:\\path\to\workbook.xlsx_
            module (str): _module name_
            macro (str): _macro (sub or function) name_
            readonly (bool, optional): _does not save the workbook after completed, Defaults to True.
            password (str, optional): _description_. Defaults to None.
        """

        if self._is_alive():
            msg = "Only one macro can be executed at a time."
            print(msg)
            return
        
        if not Path(path).exists(): return
        self.open(str(path), readonly, password) # _open_encrypted from open
        
        try:
            self.app.Application.Run(f"{path.name}!{module}.{macro}")
            print(f"{module}.{macro} Completed.")
        
            if not readonly: # might be uncessary
                self.app.Workbooks[path.name].Save()
                print(f"{path.name} Saved.")
            
            self.app.Workbooks[path.name].Close()
            print(f"{path.name} Closed.")
        
        except Exception as e:
            traceback.print_exception(e)

    def exec_macro_threaded(self,path: Union[str,Path], module:str, macro: str,
                  readonly: bool = True, password: str = None):
        """_Run is on a thread assigned to self. Only one macro execution is allowed
            whether or not it's threaded. If excel needs to be closed after macro exec,
            it must be done from within the vba module macro._

        Args:
            path (str, Path): _C:\\path\to\workbook.xlsx_
            module (str): _module name_
            macro (str): _macro (sub or function) name_
            readonly (bool, optional): _does not save the workbook after completed, Defaults to True.
            password (str, optional): _description_. Defaults to None.
        """
        
        if self._is_alive():
            msg = "Only one macro can be executed at a time."
            print(msg)
            return

        if not Path(path).exists(): return

        def exec_macro(*args):
            msg = "Excel will not close after threaded macro execution."
            print(msg)
            pythoncom.CoInitialize()
            self.open(str(path), readonly, password) 
            self.app.Run(f"{path.name}!{module}.{macro}")

        self.thread = Thread(target=exec_macro)
        self.thread.start()