import inspect, os
from time import sleep
from datetime import datetime

import pandas as pd

from process.shared import Console, Table, Notify, WindowManager, Browser


class Test_Table:

    TABLE = Table
    DF = pd.DataFrame([
        [1, 1, 1, 1, 1],
        [2, 2, 2, 2, 2],
        [3, 3, 3, 3, 3],
        [4, 4, 4, 4, 4],
        [5, 5, 5, 5, 5]], columns=["one","two","three","four","five"])

    def test_add_column(self):
        inspect.stack()[0].function
        # Also tests replace_column
        df = self.TABLE.add_column(
                self.DF.copy(),
                header="six",
                value=[6,6,6,6,6],
                colIdx=2,
                after=True,
                overwrite=True)
        print(df)
        result = "OK" if df.columns[3] == "six" else "FAILED!"
        print(result)

    def test_add_row(self):
        inspect.stack()[0].function
        df = self.TABLE.add_row(
                self.DF.copy(),
                ["x","x","x","x","x"],
                rowIdx=2,
                after=False)
        print(df)
        result = "OK" if len(df) > len(self.DF) else "FAILED!"
        print(result)

    def test_replace_row(self):
        inspect.stack()[0].function
        # [TODO] Pick one approach or the other for add/replace row/column.
        df = self.TABLE.replace_row(
                self.DF.copy(),
                ["one","two","three","four","five"],
                rowIdx=0)
        print(df)
        result = "OK" if len(df) == len(self.df) else "FAILED!"        
        print(result)

    def test_change_value(self):
        inspect.stack()[0].function
        df = self.TABLE.change_value(self.DF.copy(),1,1,"two")
        print(df)
        result = "OK" if df[df.columns[1]][1] == "two" else "FAILED!"
        print(result)

    def test_filter_rows(self):
        inspect.stack()[0].function
        df = self.TABLE.filter_rows(self.DF.copy(),[1, 2, 3],out=True)
        print(df)
        result = "OK" if len(df) == 2 else "FAILED!"

    def test_filter_columns(self):
        inspect.stack()[0].function
        df = self.TABLE.filter_columns(self.DF.copy(),["one", "two"],out=False)
        print(df)
        result = "OK" if len(df.columns) == 2 else "FAILED!"
        print(result)

    def test_reorder_columns(self):
        inspect.stack()[0].function
        df = self.TABLE.reorder_columns(self.DF.copy(),["two","five","one"])
        print(df)
        result = "OK" if df.columns == ["two","five","one"] else "FAILED!"
        print(result)

    def main(self):
        self.test_add_column()
        self.test_add_row()
        self.test_replace_row()
        self.test_change_value()
        self.test_filter_rows()
        self.test_filter_columns()
        self.test_reorder_columns()


class Test_Console:

    
    CONSOLE = Console() # tests for Table are rolled into Console.
    OPTIONS = {
        "N": "ext",
        "P": "revious",
        "<": "Go Back",
         1 : "Option One",
        "#": "Page Number?",
         "": "x or <Enter> to Exit"}

    
    def test_get_selection(self):
        inspect.stack[0][4]
        res = self.CONSOLE.get_selection(["Yes", "No"])
        result = "OK" if res == "Yes" else "FAILED!"
        print(result)

    def test_get_num_selection(self):
        inspect.stack[0][3]
        print("Selection requests can be filtered.")
        sleep(2)
        selection_filter = [1, 2, 3, 4]
        print("get_num_selection does not clear the screen. Please press 1, 2, 3, or 4 then <Enter> to continue.")
        res = self.CONSOLE.get_num_selection(selection_filter)
        result = "OK" if res in selection_filter else "FAILED!"
        print(f"Test result: {result}")

    def test_get_option_selection(self):
        inspect.stack[0][3]
        # The return value is ensured to be the same type as the options value.
        res = self.CONSOLE.get_option_selection(self.OPTIONS)
        result = "OK" if self.OPTIONS[res] else "FAILED!"
        print(result)
    
    def test_get_datetime(self):
        inspect.stack[0][3]
        # This test combines get_date, get_time, and _get_datetime.
        # The date and time components are independently parsed to
        # a verifiabled date prior to return. get_datetime returns
        # either a datetime.datetime value or a string formatted 
        # human-readable datetime.
        res = self.CONSOLE.get_datetime(asDt=True)
        result = "OK" if type(res) == datetime else "FAILED!"
        print(result)

    def test_request_table(self):
        inspect.stack[0][3]
        print("Please place a few entries into the table editor for the next test.")
        sleep(3)
        df =  self.CONSOLE.request_table(["First Name, Last Name"])
        print(df)
        result = "OK" if type(df) == pd.DataFrame else "FAILED!"
        print(result)

    def test_page_navigate(self):
        inspect.stack[0][3]
        # df is still assigned to self.CONSOLE
        df = self.CONSOLE.df
        matrix = Table.dataframe2matrix(df)
        row = matrix[0]
        matrix = [row] * 200
        df = pd.DataFrame(matrix, columns=df.columns)
        self.CONSOLE.df = df
        self.CONSOLE.page_navigate()
        print("Result: OK")

    def test_insert_row(self):
        inspect.stack[0][3]
        numRows = len(self.CONSOLE.df)
        self.CONSOLE.insert_row(after=True)
        df = self.CONSOLE.df
        print(df.head())
        result = "OK" if len(self.CONSOLE.df) > numRows else "FAILED!"
        print(result)
    
    def test_append_row(self):
        inspect.stack[0][3]
        numRows = len(self.CONSOLE.df)
        self.CONSOLE.append_row()
        df = self.CONSOLE.df
        print(df.tail())
        result = "OK" if len(self.CONSOLE.df) > numRows else "FAILED!"
        print(result)

    def test_delete_rows(self):
        inspect.stack[0][3]
        numRows = len(self.CONSOLE.df)
        self.CONSOLE.delete_rows()
        result = "OK" if len(self.CONSOLE.df) < numRows else "FAILED!"
        print(result)
    
    def delete_columns(self):
        inspect.stack[0][3]
        numCols = len(self.CONSOLE.df.columns)
        self.CONSOLE.delete_columns()
        result = "OK" if len(self.CONSOLE.df.columns) < numCols else "FAILED!"
        print(result)

    def main(self):
        self.test_get_num_selection()
        self.test_get_option_selection()
        self.test_get_datetime()
        self.test_request_table()
        self.test_page_navigate()
        self.test_insert_row()
        self.test_append_row()
        self.test_delete_rows()
        self.delete_columns()


class Test_Notify:

    
    NOTIFY    = Notify()
    _4MESSAGE = "This is a msgbox alert with a 'inform' icon. Two other icons are available: 'warn' & 'stop'."
    _4TASKBAR = "This is a less intrusive taskbar notification. They remain until cleared."

    
    def test_message(self):
        self.NOTIFY.message(self._4MESSAGE)

    def test_taskbar(self):
        self.NOTIFY.taskbar(self._4TASKBAR)

    def main(self):
        self.test_message()
        self.test_taskbar()


class Test_WindowManager:


    WINDOW = WindowManager()
        
    
    def _nxt(self):
        sleep(2)
        
    def Test_getWindowTitles(self):
        print(inspect.stack()[0].function)
        titles = WindowManager.getWindowTitles()
        result = "OK" if type(titles) == list and titles else "FAILED!"
        print(result)
        self._nxt()

    def Test_findWindowTitles(self):
        print(inspect.stack()[0].function)
        titles = WindowManager.findWindowTitles("notepad")
        result = "OK" if type(titles) == list and titles else "FAILED!"
        print(result)
        print("Please visually inspect the results from the following tests.")
        self._nxt()

    def Test_showWindow(self):
        for state in self.WINDOW.SW_STATES:
            print(f"Testing_{state}")
            self.WINDOW._showWindow(state)
            self._nxt()
        
    def Test_keepOnTop(self):
        print(inspect.stack()[0].function)
        self.WINDOW.keepOnTop()
        self._nxt()
        
    def Test_resize(self):
        print(inspect.stack()[0].function)
        self.WINDOW.move(300, 600)
        self._nxt()

    def Test_moveCenter(self):
        print(inspect.stack()[0].function)
        self.WINDOW.moveCenter()
        self._nxt()

    def Test_moveLeft(self):
        print(inspect.stack()[0].function)
        self.WINDOW.moveLeft()
        self._nxt()

    def Test_moveRight(self):
        print(inspect.stack()[0].function)
        self.WINDOW.moveRight()
        self._nxt()

    def Test_moveTop(self):
        print(inspect.stack()[0].function)
        self.WINDOW.moveTop()
        self._nxt()

    def Test_close(self):
        print(inspect.stack()[0].function)
        self.WINDOW.close()

    def main(self):
        self.WINDOW.title = "notepad"
        self.Test_findWindowTitles()
        self.Test_getWindowTitles()
        self.Test_moveCenter()
        self.Test_showWindow()
        self.Test_moveRight()
        self.Test_keepOnTop()
        self.Test_moveTop()
        self.Test_resize()
        self.Test_close()


class Test_Browser:


    WEBBROWSER = Browser()


    def test_open_as_app(self):
        print("Opening SharePoint...")
        self.WEBBROWSER.open_as_app("sharepoint")

    def test_open_as_reference(self):
        print("Opening Workday. Workday will remain on top of all other windows until closed.")
        self.WEBBROWSER.open_as_reference("workday")
        print("This concludes tests for process.shared.")

    def main(self):
        self.test_open_as_app()
        self.test_open_as_reference()



if __name__ == "__main__":

    test = Test_Table()
    test.main()

    test = Test_Console()
    test.main()

    test = Test_Notify()
    test.main()

    os.system("notpad.exe")
    res = input("Tests for Window Manager. Press <ENTER> after Notepad opens to Continue: ")
    test = Test_WindowManager()
    test.main()

    test = Test_Browser()
    test.main()