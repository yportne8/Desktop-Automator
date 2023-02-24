import traceback
from pathlib import Path
from typing import Union

import pyodbc

import pandas as pd
from win32com import client

from .shared import FileManager


# [!] This file is not complete and has not been tested.


class FileNotFound(Exception):
    pass


class FileExists(Exception):
    pass


class Access:
    
    
    def __init__(self):
        self.filemanager = FileManager()
        self._file = None
    
    def __del__(self):
        app = self.app
        app.Application.Quit()
        print("Access has closed.")

    @property
    def app(self):
        return client.Dispatch('Access.Application')
    
    @property
    def file(self):
        return self._file
    
    @file.setter
    def file(self, value):
        file = Path(value)
        
        if not file.exists():
            raise FileNotFound
        
        self._file = str(file)
        
    def close(self):
        app = self.app
        app.CloseCurrentDatabase(self.path)
    
    def open(self):
        app = self.app
        app.OpenCurrentDatabase(self.path)
    
    def get(self, name: str, format, objectType, path: Union[str, Path],
            auto: bool = False, overwrite: bool = True):
        # [NOTE] Verifies that the requested report path does not already exist.
        if self.filemanager.doesPathExist(path, overwrite):
            raise FileExists
        else:
            app = self.app
            try:
                app.DoCmd.OutputTo(
                    ObjectType=objectType, ObjectName=name,
                    OutputFormat=format, OutputFile=path,
                    AutoStart=auto)
            except Exception as e: 
                traceback.print_exception(e)


class DataBase:

    def __init__(self, path: Union[str, Path]):
        self._path = path
        self._conn = None # sets class properties
    
    def __del__(self):
        self.conn.close()
    
    @property
    def path(self):
        return self._path

    @path.setter
    def path(self, value):
        path = Path(value)
        if not path.exists(): raise FileNotFound
        self._path = path
    
    @property 
    def params(self):
        p = "DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=%s"
        return p % self.path
        
    @property
    def conn(self):
        return self._conn
    
    @conn.setter
    def conn(self, value):
        try:
            self._conn.close()
        except:
            pass
        
        if not value:
            self._conn = pyodbc.connect(self.params)
        else:
            self._conn = value
    
    @property
    def cursor(self):
        try:
            return self.conn.cursor()
        except Exception as e:
            print("Please reset the connection and try again.", f"{e}\033[0K\r")
    
    def _lst2sqlStr(self, lst: list):
        sqlStr = "("
        for item in lst[:-1]:
            sqlStr += f"{item}, "
        sqlStr += f"{lst[-1]})"
        return sqlStr
    
    def select(self, tableName: str, sqlQuery: str,
               xlPath: Union[str, Path] = None,
               csvPath: Union[str, Path] = None) -> pd.DataFrame:
        # [TODO] Sql statement parser/builder.
        query = f"SELECT FROM {tableName} " + sqlQuery.strip()
        self.cursor.execute(query)
        
        results = self.cursor.fetchall()
        df = pd.DataFrame(results)
        if xlPath:
            df.to_excel(xlPath)
        if csvPath:
            df.to_csv(csvPath)
        self.cursor.close()
    
    def insert(self, tableName: str, fields: list, values: list):
        query = f"INSERT INTO {tableName} \
                    {self._lst2sqlStr(fields)} VALUES \
                        {self._lst2sqlStr(values)}"
        values = (values)
        self.cursor.execute(query, values)
        
Access = lambda: print("This feature has not yet been released.")
Database = lambda: print("This feature has not yet been released.")