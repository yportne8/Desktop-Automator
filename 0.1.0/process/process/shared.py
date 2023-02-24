import os
import webbrowser
from time import sleep
from pathlib import Path
from pprint import pprint
from threading import Thread
from typing import Union, Any
from urllib.parse import quote_plus
from datetime import datetime as dt
from dateutil.parser import parse as dtparse

import win32com
from ctypes import wintypes, windll, byref
from win32api import GetModuleHandle, GetSystemMetrics
from win32gui import (DestroyWindow, Shell_NotifyIcon,
    UnregisterClass, CreateWindow, WNDCLASS, 
    RegisterClass, UpdateWindow, LoadImage,
    NIF_ICON, NIF_MESSAGE, NIM_ADD, NIM_MODIFY,
    NIF_TIP, NIF_INFO, ShowWindow, SetWindowPos,
    SendMessage, PostMessage, GetWindowText, 
    GetForegroundWindow, ShowWindow, 
    IsWindowVisible, GetWindowRect,
    FindWindow, EnumWindows)

import pandas as pd


class WindowNotFound(Exception):
    """
    _Used for testing._
    """
    pass


class DtFrmtStrMistyped(Exception):
    """
    _Easy to miss typing error_
    """
    pass


class DataFrameNotAssigned(Exception):
    """
    _app.df@property must be assigned prior to table operations_
    """
    pass


class RowCountMismatchValue4Df(Exception):
    """
    _User is trying to add a column of the wrong size._
    """
    pass


class UnknownlistDf4pagenavigate(Exception):
    """
    _The listDf parameter is not iterable._
    """
    pass


class MismatchRowLengthColNum(Exception):
    """
    """
    pass


class Console:
    """
    _Contains static components for fetching verified inputs from the user._
    """

    def clear():
        os.system("cls || clear")

    def print_centered_to_screen(msg: str):
        width = os.get_terminal_size.columns
        print(msg.center(width))

    def get_selection(msg, filter: list=None):
        """_Try/Abort, etc. Clears the screen._

        Args:
            filter (list, optional): _possible entries to filter the re_. Defaults to [].

        Returns:
            _(any || list)_: _the item selected in its original form or a list of the same_
        """

        Console.clear()
        
        while True:
            
            try:
                print(f"Options: {filter}")
                sel = input("Selection?: ")
                assert sel in filter
                return sel
            
            except:
                print("???\033[0K\r")


    def get_row_selection(filter: list=None):
        """_Usage includes bulk-fetching row numbers or column names from the user.
            
            Does not clear the screen._

        Args:
            filter (list, optional): _possible entries to filter the re_. Defaults to [].

        Returns:
            _(any || list)_: _the item selected in its original form or a list of the same_
        """

        try:
            assert [int(f) for f in filter]
        except:
            print("Only numeric values are allowed in the filters list.")
            return

        while True:
            
            try:
            
                res = input("Selection? [#, #-#, #,#,#]: ").strip()
            
                if "-" in res and "," in res:
                    raise ValueError
            
                elif not "-" in res and not "," in res:
                    if filter and not int(res) in filter:
                        print("Not among the filter options.")
                    else:
                        return int(res)
                
                elif "-" in res:
                    start, end = (res.split("-"))
                    start, end = start.strip(), end.strip()
                    rng = list(range(int(start), int(end)+1))
                    
                    if filter:
                        filteredRng = [i for i in rng if not i in filter]
                    
                        if len(filteredRng) != len(rng):
                            print("At least some choices are not among the filter options.")
                    
                        else:
                            return rng
                
                elif "," in res:
                    selections = res.split(",")
                    selections = [int(i.strip()) for i in selections]
                    
                    if filter:
                    
                        filteredSelections = [i for i in rng if not i in filter]
                    
                        if len(filteredSelections) != len(selections):
                            print("At least some choices are not among the filter options.")
                    
                        else:
                            return selections
            
            except:
                print("???\033[0K\r")
                sleep(2)

    def get_option_selection(msg: str, options: dict,
                             accept_numbers: bool=False,
                             accept_enter: bool=False) -> Any:
        """
        _Gets the user selection from a choice of options. Returns the
         value of the option key, not the option._
         Each option is printed on a separate row followed by a selection
         request. The user cannot leave the selection screen until one
         of the possible selections (options.keys) are selected. The 
         return value is in the same type as that of the key, unless
         the accept_numbers parameter is set to True, as would be
         necessary in the case of a multipage menu:
         
         [N]ext
         [B]ack
         [#]Page
         
         Selection? [?]:_
         
         In this case, the calling function should contain additional
         checks to verify that the user's response is acceptable._

        Args:
            msg (str): _message to user_
            options (dict): _key-to-press: option_

        Returns:
            obj: _returns the key in the type received_
        """
        
        Console.clear()
        for oKey, o in options.items(): pprint(f"[{oKey}] {o}")
        retValues = {str(oKey).lower(): o for oKey, o in options}
        optionKeys = list(options.keys())
        
        while True:
            res = input(f"{msg} [?]: ").strip().lower()
            try:
                if accept_enter and not res:
                    return res
                elif accept_numbers and res.isnumeric():
                    return int(res)
                elif accept_numbers and not res.isnumeric():
                    raise ValueError
                else:
                    idx = retValues.index(res)
                    return optionKeys[idx]
            except:
                print("???\033[0K\r")
                sleep(2)

    def get_datetime(title: str, asDt: bool = True) -> Union[dt, str]:
        """_gets a date and time from the user either as a string or as datetime.datime.
            date and time are parsed to a verifiable date and time by the intermediary 
            get_date and get_time functions. Collection for the the date and time are
            separate. title is used to identify the purpose for the date and time 
            collection to the user_

        Args:
            msg (str): _message to the user_
            asDt (bool, optional): _as datetime.datetime_. Defaults to True.

        Raises:
            DtFrmtStrMistyped: _the format string for datetime.datetime to str is mistyped_

        Returns:
            _dt | str_: _datetime.datetime or str based on the asDt parameter._
        """

        pprint(title) # date and time requested separately
        msg = "Please enter the date"
        date = Console.get_date(msg)
        datetime = f"{date}"

        msg = "Please enter the time"
        hour, min = Console.get_time(msg)
        
        hour = f"0{hour}" if hour < 9 else hour
        min = f"0{min}" if min < 9 else min
        datetime += f" {hour}:{min}"
        
        if asDt:
            try:
                return dtparse(datetime)
            except:
                raise DtFrmtStrMistyped 
        else:
            return datetime

    def get_time(msg: str) -> tuple:
        """
        _Gets user input parsed to a verifiable time of day._

        Returns:
            _tuple_: _(hour, minutes)_
        """
        Console.clear()
        while True:
            try:
                hours = input(f"{msg} [Hour]: ").strip()
                mins = input(f"{msg} [Minutes]: ").strip()
                _ = dtparse(f"{dt.today().strftime('%m/%d/%Y')} {hours}:{mins}")
                return (int(hours), int(mins))
            except:
                print("???\033[0K\r") # auto erase flash temp error message
                sleep(1)

    def get_date(msg: str) -> str:
        """
        _Gets user input parsed to a verifiable date._

        Returns:
            _type_: _description_
        """
        Console.clear()
        while True:
            try:
                date = input(f"{msg} [mm/dd/yyyy]: ")
                date = dtparse(date) # verify that the date is parsable.
                date = date.strftime("%m/%d/%Y")
                return date
            except Exception as e:
                print("???\033[0K\r")
                sleep(2)

    def get_yesno(msg: str, default: str="yes") -> str:
        """
        _Gets a a/b choice from the user, does not clear the screen 
         in case information has been pre-printed_

        Args:
            msg (str): _message to user_
            defaultYes (bool, optional): _is the default option yes_. Defaults to True.

        Returns:
            str: _yes or no_
        """
        
        msg = f"{msg} [Y/n]: " if default else f"{msg} [y/N]: " 
        
        res = input(msg).strip()
        if res and res == "n":
            return "no"
        else:
            return "yes"


class Table(Console):
    """
    _Provides table input on the console.
     Table can be imported and used separately, all functions are static._
    """

    def clear():
        """
        _Clears the system prompt._
        """

        os.system('cls||clear')


    def get_table_row(headers: list):
        """_summary_

        Args:
            headers (list): _description_

        Returns:
            _type_: _description_
        """

        while True:
            Console.clear()
            print(f"Columns: {headers}") # column heading reference for the user
            
            row = input(f"[, , , <Enter> to Verify]: ").strip().split(",")
            row = [col.strip() for col in row] # cleanup any extra spaces
            
            if len(headers) != len(row): 
                # asymmetric matrix
                msg = "[!] There was a mismatch in the number of columns provided."
                pprint(msg)
                res = input("Type '!' for details or just <Enter> to try again: ").strip()
                if res == "!":
                    msg = "The number of columns in the printed Columns Headers must match the number " & \
                          "of commas provided, even if some commas are followed by empty place holders: , ,"
                    pprint(msg)
                    _ = input("<Enter> to continue: ")
                # returns to top of while loop from here.
            else:
                return row

    def dataframe2matrix(df: pd.DataFrame):
        """_Converts a dataframe into a nested matrix._

        Args:
            df (pd.DataFrame): _pandas.DataFrame_

        Returns:
            _list_: _converted list_
        """

        data = list()
        # [NOTE] pd.DataFrame does not have a to_list option.
        for idx in df.index:
            # row and columns have a to_list option.
            data.append(df.loc[idx].to_list())
        return data

    def add_column(df: pd.DataFrame, header: Union[str,int],
                   value: Any, colIdx: int=None,
                   after: bool=True, overwrite: bool=False) -> pd.DataFrame:
        """_Add a column to a dataframe._

        Args:
            header (str): _column header_
            value (Any): __
            df (pd.DataFrame): _description_

        Returns:
            pd.DataFrame: _description_
        """

        try:
            if not overwrite and header in df.columns:
                print("To overwrite an existing column, \
                      please change the overwrite parameter to True.")
                return

            if type(value) in [str,int,float]:
                value = [value] * len(df)
            elif type(value) == dict():
                # [NOTE] all components of a dataframe will 
                # raise an exception if place in a if/and.
                idx = df.index.to_list()
                try:
                    assert list(value.keys()) == idx
                    value = list(value.values())
                except:
                    rearrangedValue = list()
                    for idx in df.index:
                        rearrangedValue.append(value[idx])
                    value = rearrangedValue
            
            df[header] = value
            
            if colIdx:
                colIdx = colIdx+1 if after else colIdx    
                columns = df.columns[:-1]
                columns = columns[:colIdx]+[header]+columns[colIdx:]
                df = df[[(columns)]].copy()
            return df                
        except:
            RowCountMismatchValue4Df

    def add_row(df: pd.DataFrame, value: list, rowIdx: int=None, after: bool=True) -> pd.DataFrame:
        """_Add a row into a dataframe._

        Args:
            df (pd.DataFrame): _pandas.DataFrame_
            rowIdx (int, optional): _row number for insert_. Defaults to None.
            after (bool, optional): _insert after the row_. Defaults to True.

        Returns:
            pd.DataFrame: _description_
        """
        dfrow = pd.DataFrame([value], columns=df.columns)
        if rowIdx:
            rowIdx = rowIdx + 1 if after else rowIdx
            if not rowIdx >= len(df):
                dfs = [df[:rowIdx].copy(), dfrow, df[rowIdx:].copy()]
            else:
                dfs = [df, dfrow]    
        else:
            dfs = [df, dfrow]
        return pd.concat(dfs)

    def replace_column(df: pd.DataFrame, header: str,
                       value: Any, colIdx: int=None,
                       after: bool=True) -> pd.DataFrame:
        """_Replaces a column of the dataframe._

        Args:
            df (pd.DataFrame): _pandas.DataFrame_
            header (str): _column header_
            value (Any): _value(s) to place into column_
            colIdx (int, optional): _column index_. Defaults to None.
            after (bool, optional): _insert after the index column_. Defaults to True.

        Returns:
            pd.DataFrame: _pandas.DataFrame_
        """
        
        return Table.add_column(df,header,value,colIdx,after,True)

    def replace_row(df: pd.DataFrame, value: list, rowIdx: int=None) -> pd.DataFrame:
        """_Replace a row in the dataframe._

        Args:
            df (pd.DataFrame): _pandas.DataFrame_
            value (list): _column header_
            rowIdx (int, optional): _row index_. Defaults to None.

        Returns:
            pd.DataFrame: _description_
        """
        
        if len(value) != len(df.columns):
            print(f"The length of the value parameter must match the length of df.columns: {len(df.columns)}")
        else:
            dfreplacement = pd.DataFrame([value], columns=df.columns)
            return pd.concat([df[:rowIdx], dfreplacement, df[rowIdx+1:]])

    def change_value(df: pd.DataFrame, column: Union[str, int], row: int, value: Any):
        """_Changes a value of a single cell in the assigned table_

        Args:
            columnName (str): _name of the column_
            rowIdx (int): __
            newValue (Any): _description_
        """

        if type(column) == int:
            column = df.columns[column]
        
        try:
            df[column][row] = value
            return df
        except:
            print("The table value could not changed.")
        

    def filter_rows(df: pd.DataFrame, selections: list, out: bool=True):
        """_Filters the rows of a dataframe_

        Args:
            out (bool, optional): _whether to filter the selected rows out_. Defaults to True.
        """
        print("Warning: Filtered tables cannot be unfiltered.")
        if type(selections) == int:
            return df[selections].copy
        else:
            headers = df.columns
            matrix = Table.dataframe2matrix(df)
            matrix4filteredDf = list()
        
            for row in matrix:
                if out:
                    if not row in selections:
                        matrix4filteredDf.append(row)
                else:
                    if row in selections:
                        matrix4filteredDf.append(row)
            return pd.DataFrame(matrix, columns=headers)

    def filter_columns(df: pd.DataFrame, selections: list, out: bool=True):
        """_Filters the columns of a dataFrame_

        Args:
            df (_type_): _description_
            out (bool, optional): _description_. Defaults to True.

        Returns:
            _type_: _description_
        """
        if out:
            columns = list()
            for i, column in df.columns:
                if not i in selections:
                    columns.append(column)    
        else:
            columns = [df.columns[i] for i in selections]
        return df[[(columns)]].copy()

    def reorder_columns(df: pd.DataFrame, columns: list):
        """
        _Reorders the columns of a dataframe.
            
         The 'columns' parameter should reflect the columns
         to be retained and the order in which to retain them._

        Args:
            df (pd.DataFrame): _pd.DataFrame_
            columns (list): _of reordered columns_

        Returns:
            _type_: _A dataframe with only the reordered columns._
        """
    
        return df[[(columns)]].copy()


class Console(Console):
    """
    _Provides verified data collection on the console._

    Raises:
        DtFrmtStrMistyped: _format string for dt.strftime_
    """

    def __init__(self):
        """_The editor does not need to be initialized as it houses static table operation functions.
            The assignment of an empty dataframe to the self.df property that can be populated entiredly
            from the console and exported to file._
        """

        self._df = pd.DataFrame() # empty worksheet
        self.table = Table # excel for dataframes on console

    @property
    def df(self) -> pd.DataFrame:
        # A copy is operated on, then the property reassigned.
        return self._df.copy()

    @df.setter
    def df(self, value: pd.DataFrame):
        self._df = value

    def request_table(self, headers: Union[str, list]) -> df:
        """
        _For loading small sets of data manually.
         Creates a new dataframe via input from the command line.
         As with all other dataframe/table operations performed by Console,
         the generated table is centrally stored in self.df._
        """
        if type(headers)==str: headers = [c.strip() for c in headers.split(',')]
        table = list()

        while True: 
            row = self.table.get_table_row(headers=self.df.columns)
            if row:
                table.append(row)
            
            else:
                break
        
        self.df = pd.DataFrame(table, columns=headers)
        print(".df has been assigned")

    def page_navigate(self, linesPerScreen: int=50, page=0):
        """_Multipage navigation of self.df._

        Args:
            linesPerScreen (int, optional): _lines per screen_. Defaults to 50.
            page (int, optional): _page_. Defaults to 0.

        Returns:
            _type_: _description_
        """

        startIdx = page * linesPerScreen
        try:
            df = self.df
            print(df[startIdx:linesPerScreen])
            menu = {
                "#": "Go to Page",
                ">": "Next Page",
                "<": "Previous Page",
                "" : "Press Enter to exit"}

            res = self.table.get_option_selection(options=menu)
            if res == "#":
                print("A number is required for page selection.")
            elif type(res) == int:
                page = res
            elif res == ">":
                page += 1
            elif res == "<":
                page -= 1
            return self.page_navigate(df,linesPerScreen,page)
        
        except:
            if page == 0:
                raise UnknownlistDf4pagenavigate
            else:
                page = 0
                return self.page_navigate(df,linesPerScreen,page)
  
    def insert_row(self, after: bool=True):
        """
        _Inserts a row into the assigned table._

        Args:
            after (bool, optional): _description_. Defaults to True.
        """

        self.table.clear()
        df = self.df
        print(df.head())
        while True:
            try:
                idx = input("Index Number?: ")
                assert type(idx) == int

                row = self.table.get_table_row(df.columns)
                self.df = self.table.add_row(df,row,idx,after)
                break      
            except:
                print("The insert index should be one number.")
                sleep(2)
  
    def append_row(self):
        """
        _Appends a row into the assigned table._
        """

        df = self.df
        row = self.table.get_table_row(df.columns)
        self.df = pd.concat(
                    [df, pd.DataFrame([row],
                    columns=df.columns)])

    def delete_rows(self):
        print(f"Please provide row numbers to filter out?")
        selections = Table.get_selection()
        
        print("Warning: Row deletions are permanent.")
        self.df = Table.filter_rows(self.df, selections, out=True)

    def delete_columns(self):
        print("Warning: Filtered tables cannot be unfiltered.")
        df = self.df.copy()
        
        for i, column in enumerate(df.columns):
            print(f"[{i+1}] {column}")
        
        selections = Table.get_selection(list(range(1,len(df)-1)))
        self.df = Table.filter_columns(self.df, selections, out=True)


class Notify:
    
    """
    _Allows for native messagebox and taskbar notifications with various options._
    """
    
    ICO = str(Path(Path(__file__).parent, "notifier.ico"))
    BTNOPTIONS = ["ok","yesno","okcancel","retrycancel"]
    FONT = ("Segoe UI", 11)
    PADXY = "1m"      

    def _notify(self, msg, title):
        """
        _Taskbar icon notification._

        Args:
            msg (_type_): _description_
            title (_type_): _description_
        """
        
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = f"{title} Notifier"
        classAtom = RegisterClass(wc)
        style = win32com.WS_OVERLAPPED | win32com.WS_SYSMENU
        hwnd = CreateWindow(
                classAtom, "Taskbar", style,
                0, 0, win32com.CW_USEDEFAULT, win32com.CW_USEDEFAULT,
                0, 0, hinst, None)
        UpdateWindow(hwnd)
        
        icon_flags = win32com.LR_LOADFROMFILE | win32com.LR_DEFAULTSIZE
        hicon = LoadImage(hinst, self.ICO, win32com.IMAGE_ICON, 0, 0, icon_flags)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (hwnd, 0, flags, win32com.WM_USER+20, hicon, "tooltip")
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(
            NIM_MODIFY, 
            (hwnd, 0, NIF_INFO, win32com.WM_USER+20,
             hicon, "Notification", msg, 200, title))
        sleep(3)
        DestroyWindow(hwnd)
        UnregisterClass(classAtom, hinst)
        
    def message(self, msg: str, icon: str="inform", title: str="DXC Office") -> str:
        """
        _Uses a native message box to alert the user of monitored changes.
         
         The following icons are available:
         
         : warn :
         : stop :
         : inform :
         
         The title defaults to "DXC Office" but can be changed as needed._

        Args:
            msg (_str_): _message to user_
            icon (_str_): _warn, stop, inform_
            title (str, optional): _title for msgbox_. Defaults to "DXC Office".
        """

        icons = {
            "warn": 48,
            "stop": 16,
            "inform": 64}
        _ = windll.user32.MessageBoxW(0,msg,title,0|icons[icon]|0)
               
    def taskbar(self, msg: str, title: str="DXC Office"):
        """_Triggers a taskbar notification. The notification title
            must state the name of the calling program. However, the
            title line, if provided, is printed in larger print above
            the notification message._

        Args:
            msg (str): _message to user_
            title (str, optional): _title for msgbox_. Defaults to "".
        """

        if Path(self.ICO).exists():
            thrd = Thread(target=self._notify, args=(msg, title))
            thrd.start()
            sleep(3)
            thrd.join()
        else:
            msg = f"Error: {self.ICO} is required for the notify function. \
                    Please return or replace this file."
            print(msg)


class Keyboard:
    """ 
    _Simulates Keyboard strokes onto a Windows window. All functions are static,
     however, they require a window hwnd as a parameter. The Class also stores a
     map of keyboard keys mapping readable text identifiers and corresponding 
     win32 parameter values._
    """
    
    WM_KEYDOWN = 256
    WM_KEYUP = 257
    KEYMAP = {
        'LBUTTON': 1, 'RBUTTON': 2, 'MBUTTON': 4,
        'XBUTTON1': 5, 'XBUTTON2': 6, 'CANCEL': 3,
        'BACK': 8, 'TAB': 9, 'CLEAR': 12, 'RETURN': 13,
        'SHIFT': 16, 'CONTROL': 17, 'MENU': 18, 'PAUSE': 19,
        'CAPITAL': 20, 'ESCAPE': 27, 'SPACE': 32, 'PRIOR': 33,
        'NEXT': 34, 'END': 35, 'HOME': 36, 'LEFT': 37, 'UP': 85,
        'RIGHT': 39, 'DOWN': 40, 'SELECT': 41, 'PRINT': 42,
        'EXECUTE': 43, 'SNAPSHOT': 44,'INSERT': 45, 'DELETE': 46,
        'HELP': 47, '0': 48, '1': 49, '2': 50, '3': 51, '4': 52,
        '5': 53, '6': 54, '7': 55, '8': 56, '9': 57, 'A': 65,
        'B': 66, 'C': 67, 'D': 68, 'E': 69, 'F': 70, 'G': 71,
        'H': 72, 'I': 73, 'j': 74, 'K': 75, 'L': 76, 'M': 77,
        'N': 78, 'O': 79, 'P': 80, 'Q': 81, 'R': 82, 'S': 83,
        'T': 84, 'V': 86, 'W': 87, 'X': 88, 'Y': 89, 'Z': 90,
        'LWIN': 91, 'RWIN': 92, 'APPS': 93, 'SLEEP': 95,
        'NUMPAD0': 96, 'NUMPAD1': 97, 'NUMPAD2': 98,
        'NUMPAD3': 99, 'NUMPAD4': 100, 'NUMPAD5': 101,
        'NUMPAD6': 102, 'NUMPAD7': 103, 'NUMPAD8': 104,
        'NUMPAD9': 105, 'MULTIPLY': 106, 'ADD': 107,
        'SEPARATOR': 108, 'SUBTRACT': 109, 'DECIMAL': 110,
        'DIVIDE': 111, 'F1': 112, 'F2': 113, 'F3': 114,
        'F4': 115, 'F5': 116, 'F6': 117, 'F7': 118,
        'F8': 119, 'F9': 120, 'F10': 121, 'F11': 122,
        'F12': 123, 'F13': 124, 'F14': 125, 'F15': 126,
        'F16': 127, 'F17': 128, 'F18': 129, 'F19': 130,
        'F20': 131, 'F21': 132, 'F22': 133, 'F23': 134,
        'F24': 135, 'NUMLOCK': 144, 'SCROLL': 145,
        'LSHIFT': 160, 'RSHIFT': 161, 'LCONTROL': 162,
        'LCTRL': 162, 'RCONTROL': 163, 'RCTRL': 163, 
        'LALT': 164, 'RALT': 165, 'BROWSER_BACK': 166,
        'BROWSER_FORWARD': 167, 'BROWSER_REFRESH': 168,
        'BROWSER_STOP': 169, 'BROWSER_SEARCH': 170,
        'BROWSER_FAVORITES': 171, 'BROWSER_HOME': 172,
        'VOLUME_MUTE': 173, 'VOLUME_DOWN': 174, 'VOLUME_UP': 175,
        'MEDIA_NEXT_TRACK': 176, 'MEDIA_PREV_TRACK': 177,
        'MEDIA_STOP': 178, 'MEDIA_PLAY_PAUSE': 179,
        'LAUNCH_MAIL': 180, 'LAUNCH_MEDIA_SELECT': 181,
        'LAUNCH_APP1': 182, 'LAUNCH_APP2': 183,
        '?': 191, '~': 192, '{': 219, '|': 220, '}': 221, '"': 222,
        'ATTN': 246, 'CRSEL': 247, 'EXSEL': 248, 'PLAY': 250, 'ZOOM': 251}
        
    def hold(key: Union[str,int], hwnd: int):
        """
        key (Union[str,int]): _single Keyboard glyph_ 
        hwnd_title (Union[str,int): _window handle number or title_
        """
        
        key = Keyboard.KEYMAP[key] if type(key) == str else key
        SendMessage(hwnd, Keyboard.WM_KEYDOWN, key, 0)

    def release(key: Union[str,int], hwnd: int):
        """
        key (Union[str,int]): _single Keyboard glyph_ 
        hwnd_title (Union[str,int): _window handle number or title_
        """
        
        key = Keyboard.KEYMAP[key] if type(key) == str else key
        SendMessage(hwnd, Keyboard.WM_KEYUP, key, 0)
    
    def press(key: Union[str,int], hwnd: int):
        """
        key (Union[str,int]): _single Keyboard glyph_ 
        hwnd_title (Union[str,int): _window handle number or title_
        """
        
        key = Keyboard.KEYMAP[key] if type(key) == str else key
        # [NOTE] press and release
        Keyboard.hold(key, hwnd)
        Keyboard.release(key, hwnd)
    
    def shortcut(keys: list, hwnd: int):
        """
        key (Union[str,int]): _single Keyboard glyph_ 
        hwnd_title (Union[str,int): _window handle number or title_
        """
        
        # [NOTE] The first key is the 'down key.'
        _ = [Keyboard.hold(k, hwnd) for k in keys]
        _ = [Keyboard.release(k, hwnd) for k in keys[::-1]]

    def series(keys: list, hwnd: int):
        """
        key (Union[str,int]): _single Keyboard glyph_ 
        hwnd_title (Union[str,int): _window handle number or title_
        """
        
        for i, k in enumerate(keys):
            keys[i] = Keyboard.KEYMAP[k] if type(k) == str else k   
        _ = [Keyboard.press(k, hwnd) for k in keys]


class WindowManager:

    """ 
    _Provides external control over a window on Windows.
     Window state change options are pulled directly from
     microsoft's win32 api documentation, resulting in 
     overlapping functionality in several functions.
     Such functions follows pywin32's camelCase naming methods._
    """

    SW_STATES={
        "hide"              : 0,
        "unhide"            : 1,
        "displayActivate"   : 1,
        "showMinimized"     : 2,
        "showMaximized"     : 3,
        "showNoActivate"    : 4,
        "showActivate"      : 5,
        "minimize"          : 6,
        "minimizeNoActivate": 7,
        "showNoActivate"    : 8,
        "restore"           : 9,
        "default"           : 10,
        "minimizeThreaded"  : 11}
    HWND_TOPMOST=-1
    SHEIGHT=1
    SWIDTH=0
    SM_CYSIZEFRAME=34
    WM_CLOSE=16
    
    def __init__(self):
        """
        Args:
            title (str, optional): _partial titles accepted_. Defaults to None.
        """
        
        self.taskbarOffSet = GetSystemMetrics(self.SM_CYSIZEFRAME)
        # [NOTE] Visual estimate.
        self.hoverEffect = self.taskbarOffSet // 5 
        
        # Resolves the taskbar issue.
        workingArea = wintypes.RECT()
        _ = windll.user32.SystemParametersInfoW(48,0,byref(workingArea),0)
        self.sw = workingArea.right - workingArea.left
        self.sh = workingArea.bottom - workingArea.top
        
        # [NOTE] Assignment of title is required for class methods.
        self._title = None

    def press(self, key: Union[str,int]):
        Keyboard.press(key,self.hwnd)
    
    def hold(self, key: Union[str,int]):
        Keyboard.hold(key,self.hwnd)
    
    def release(self, key: Union[str,int]):
        Keyboard.release(key,self.hwnd)
    
    def shortcut(self, keys: list):
        Keyboard.shortcut(keys,self.hwnd)

    def series(self, keys: list):
        Keyboard.series(keys,self.hwnd)

    def displayActivate(self):
        self._showWindow("displayActivate")
          
    def showMinimized(self):
        self._showWindow("showMinimized")
        
    def showMaximized(self):
        self._showWindow("showMaximized")
        
    def showNoActivate(self):
        self._showWindow("showNoActivate")
        
    def showActivate(self):
        self._showWindow("showActivate")
        
    def minimizeNoActivate(self):
        self._showWindow("minimizeNoActivate")
        
    def showNoActivate(self):
        self._showWindow("showNoActivate")
        
    def minimizeThread(self):
        self._showWindow("minimizeThread")
        
    def restore(self):
        self._showWindow("restore")
        
    def default(self):
        self._showWindow("default")

    def hide(self):
        self._showWindow("hide")
        
    def unhide(self):
        self._showWindow("displayActivate")

    def minimize(self):
        self._showWindow("minimize")

    def maximize(self):
        self.showMaximized(self)

    def close(self):
        PostMessage(self.hwnd,self.WM_CLOSE,0,0)

    def move_top_center(self):
        self.move_center()
        self.move_top()
    
    def move_top_left(self):
        self.move_left()
        self.move_top()

    def move_top_right(self):
        self.move_right()
        self.move_top()

    @staticmethod
    def getWindowTitles() -> list:
        """
        Returns:
            _list_: _opened window titles_
        """
        
        titles = list()
        def getTitle(hwnd, _x): # [NOTE] Bug patch.
            if IsWindowVisible(hwnd):
                titles.append(titles.append(GetWindowText(hwnd)))
        EnumWindows(getTitle, None)
        return [t for t in titles if t]

    @staticmethod
    def findWindowTitles(partial_title: str) -> list:
        """
        _Returns a list of titles for visible windows matching the partial_title parameter._

        Args:
            partial_title (str): _a partial (or full) title of a visible window_
            
        Raises:
            WindowTitleNotFound: _failed to find any matching titles_

        Returns:
            _list_: _list of matching window titles_
        """
        
        titles = WindowManager.getWindowTitles()
        try:
            titles = [t for t in titles if partial_title.lower() in t.lower()]
            if not titles:
                raise WindowNotFound
            return titles
        except Exception as e:
            print("Exception",f"{e}/r")
        
    @property
    def title(self):
        """
        _Returns the window title._

        Args:
            title (_str_): _window text identifier_
        """
        
        try:
            hwndTitle = GetWindowText(self.hwnd)
            if self._title != hwndTitle:
                self._title = hwndTitle
        except:
            pass
        return self._title

    @title.setter
    def title(self, value: str):
        """
        _Sets the window title. The window title 
        is verified to exist among the list of opened 
        window titles prior to assignment._

        Args:
            value (str): _description_

        Raises:
            WindowNotFound: _description_
        """
        
        if value:
            # [NOTE] Allows for partial title assignment.
            try:
                titles = self.findWindowTitles(value)
                if len(titles) > 1:
                    print(f"Several matching titles were found:\n{titles}")
                value = titles[0]
            except:
                raise WindowNotFound

        self._title = value
        if value:
            print(f"Window title has been set to: {value}.")

    @property
    def hwnd(self):
        """_Returns the window handle._

        Raises:
            WindowNotFound: _FindWindow could not find window_

        Returns:
            _type_: _description_
        """
        # [NOTE] The intermediary step of finding the window's handle
        # ensures that the calling method is operating on a opened
        # window.
        
        try:
            hwnd = FindWindow(None, self.title)
            if not hwnd:
                raise ValueError
            return hwnd
        except:
            raise WindowNotFound

    @property
    def focused(self) -> bool:
        """
        _Is the window in the foreground._
        
        Returns:
            _bool_: _current window is in the foreground_
        """
        
        return self.title == GetWindowText(GetForegroundWindow())
        
    @property
    def visible(self) -> bool:
        """
        _Is the window visible._
        
        Returns:
            _bool_: _current window visible_
        """
        
        return IsWindowVisible(self.hwnd)

    @property
    def position(self) -> tuple:
        """
        _Returns the current position of the window._
        
        Returns:
            _tuple_: _(x, y)_
        """
        
        rect = GetWindowRect(self.hwnd)
        x = rect[0] + 7
        y = rect[1]
        return (x, y)
    
    @property
    def size(self) -> tuple:
        """
        _Returns the current size of the window._
        
        Returns:
            _tuple_: _(w, h)_
        """
        
        rect = GetWindowRect(self.hwnd)
        w = rect[2] - self.position[0] - 7
        h = rect[3] - self.position[1] - 7
        return (w, h)

    def _specs(self) -> dict:
        return {"text"    : self.title,
                "hWnd"    : self.hwnd,
                "size"    : self.size,
                "position": self.position}

    def _showWindow(self, state: str):
        ShowWindow(self.hwnd,self.SW_STATES[state])

    def keep_on_top(self):
        """
        _keep a window on top of all other windows_
        """
        
        self._showWindow("displayActivate")
        w, h = self.size
        x, y = self.position
        SetWindowPos(self.hwnd,self.HWND_TOPMOST,x,y,w,h,0)

    def move(self, x: int, y: int):
        """
        _Moves the window to x, y position on the main screen._

        Args:
            x (int): _new x-axis position_
            y (int): _new y-axis position_
        """
        
        # [NOTE] x, y for this function group are converted here.
        # This function is called by other class functions that
        # change the window's position.
        x, y = int(x), int(y)
        w, h = self.size
        SetWindowPos(self.hwnd,0,x,y,w,h,0)

    def move_center(self):
        """
        _move the window to screen center_
        """
        
        x = self.sw//2 - self.size[0]//2
        y = (self.sh - self.size[1]) // 2
        self.move(x, y)

    def move_top(self):
        """
        _move the window to the topside edge_
        """
        
        x = self.position[0]
        y = self.hoverEffect
        self.move(x, y)

    def move_left(self):
        """
        _move the window to the left side edge_
        """
        
        x, y = self.position
        x = self.hoverEffect
        self.move(x, y)
        
    def move_right(self):
        """
        _move the window to the right side edge_
        """
        
        x, y = self.position
        x = self.sw - self.size[0] - self.hoverEffect
        return self.move(x, y)
            
    def resize(self, w: int, h: int):
        """
        w (int): _window width_
        h (int): _window height_
        """
        # [NOTE] w, h for this function group are converted here.
        # This function is called by other class functions that
        # change the window's size.
        w, h = int(w), int(h)
        x, y = self.position
        SetWindowPos(self.hwnd,0,x,y,w,h,0)

    def table_reference(self):
        """
        _opens the browser as an always on top app on the bottom half of the screen_
        """
        self.resize(self.sw, self.sh//2) 
        self.move(0, self.sh//2)
        self.keep_on_top()
    

class Browser:

    """
    _Opens stored urls as an app with an option to open as a always-on-top reference window_
    """

    WEBPAGES = {
        "sharepoint": "https://dxcportal.sharepoint.com/",
        "workday": "https://uid.dxc.com/app/workday/exk6a1kbskdNn42uQ5d6/sso/saml",}

    def open_as_app(self, webpage: str):
        """
        _Opens the webpage url as a web app._

        Args:
            webpage (str): _description_
        """
        cmd = "start msedge --new-window --app=%s"
        try:
            os.popen(cmd % self.WEBPAGES[webpage])
            return True
        except:
            webbrowser.open(
                f"https://google.com/search?{quote_plus(webpage)}")
            print("Unknown url.")
            sleep(2.5)
            return False

    def open_as_reference(self, webpage: str):
        """
        _Opens the web app as a always-on-top paneled-window that positions itself
         above the taskbar on the bottom half of the screen for data entry related
         tasks._
        """
        window, title = WindowManager(), "universal id"
        if self.open_as_app(webpage=webpage):
            while True:
                try:
                    window.title = title
                    window.table_reference()
                    break
                except:
                    # waiting on browser window to open
                    pass
        else:
            try:
                window.title = "google - Google Search"
                window.table_reference()
            except:
                pass