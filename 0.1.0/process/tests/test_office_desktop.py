import shutil
import inspect
import unittest
from time import sleep
from pathlib import Path

import pandas as pd

from process.office import Outlook, Excel, Word, FileManager


class Test_desktop:

    """
    _tests for the desktop module. without framework 
     because pywin32 objs cannot be threaded by unittest._
    """
    
    password = "openme"
    testcsv = Path(Path(__file__).parent, "tests", "assets", "test.csv")
    macrobook = Path(Path(__file__).parent, "tests", "assets", "macrobook.xls")
    macrobook_encrypted = Path(Path(__file__).parent, "tests", "assets", "pw_openme_macrobook.xls")
    document = Path(Path(__file__).parent, "tests", "assets", "document.docx")
    document_encrypted = Path(Path(__file__).parent, "tests", "assets", "pw_openme_document.docx")
    csv = Path(Path(__file__).parent, "tests", "assets", "csv.csv")
    
    
    def test__get_local__as_dataframe(self):
        print(inspect.stack()[0][3])
        excel = Excel()
        df = excel._get_local__as_dataframe(self.macrobook)
        
        print(df.head())
        result = "OK" if len(df) > 1 else "FAILED"
        print(result)
        #print("testing_excel_get_dataframe_encrypted")
        #df = excel._get_local__as_dataframe(
        #self.macrobook_encrypted,1,password=self.password)
        #print(df.head())
        #result = "OK" if len(df) > 1 else "FAILED"
        #print(result)

    def test__load_dataframe(self):
        print(inspect.stack()[0][3])

        if self.testcsv.exists():
            self.testcsv.unlink()
        
        df = pd.DataFrame(["OK"], columns=["result"])
        
        excel = Excel()
        excel._load_dataframe(
            self.testcsv,df,sheet=1,
            delete_if_existing=True)
        
        df = pd.read_csv(self.testcsv)
        print(df.result[0])
    
    def test_excel_exec_macro(self):
        print(inspect.stack()[0][3])
        
        module = "TestModule"
        macro = "TestSub"
        
        excel = Excel()
        excel.exec_macro(
            path=self.macrobook,
            module=module,
            macro=macro)
        del excel # excel is closed via deletion or garabage collection.

    def test_word_open(self):
        print(inspect.stack()[0][3])
        
        word = Word()
        document = word.open(self.document)
        
        print("Hiding Word.")
        word.hide()
        sleep(2)
        print("Unhiding Word")
        word.unhide()
        sleep(1)
        
        print(f"{document.Name} closing.")
        document.Close()
        del word
        
    def test_outlook_draft_email(self):
        print(inspect.stack()[0][3])
        outlook = Outlook()

        subject="Testing save draftEmail"
        body="Results: OK"
        mail = outlook.draft_email(body,subject)
        result = "OK" if mail else "FAILED"
        if result == "OK": print(f"Draft email saved to outlook.")
        
    def test_outlook_send_email(self):
        print(inspect.stack()[0][3])

        outlook = Outlook()
        to = "ckim32@dxc.com"
        body = "Result: OK. This concludes tests for desktop."
        subject = "Testing process.office.desktop.outlook"
        
        outlook.send_email(to=to,subject=subject,body=body)

    def main(self):
        self.test__get_local__as_dataframe()
        self.test__load_dataframe()
        self.test_excel_exec_macro()
        self.test_word_open()
        self.test_outlook_draft_email()
        self.test_outlook_send_email()


class Test_FileManager(unittest.TestCase):

    """
    _Tests for FileManager_
    """
        
    extractionPath = Path(Path(__file__).parent, "tests", "extraction")
    collectionPath = Path(Path(__file__).parent, "tests", "collection")
    testFolder = Path(Path(__file__).parent, "tests", "subfolderswithfiles.zip")
    UserOneDrivePath = Path(Path.home(), "OneDrive")
    CommercialOneDrivePath = Path(Path.home(), "OneDrive - DXC Production")

    def test_extract_all(self):
        targetPath = self.extractionPath
        if targetPath.exists():
            shutil.rmtree(str(targetPath))
        FileManager().extract_all(self.testFolder, targetPath)
        self.assertTrue(targetPath.exists())

    def test_resolve_path(self):
        testPath = Path(Path.home(), "Documents")
        print(f"Test Path: {str(testPath)}")
        result = str(FileManager().resolve_path(testPath))
        print(f"Result Path: {result}")
        self.assertTrue(result)

    def test_rename_existing_path(self):
        print(self.extractionPath)
        newPath = FileManager().rename_existing_path(self.extractionPath)
        Path(newPath)
        self.assertTrue(newPath!=self.extractionPath)
    
    def test_localpath_properties(self):
        desktop = FileManager().desktop
        documents = FileManager().documents
        self.assertTrue(desktop.exists() and documents.exists())
       

            
if __name__ == "__main__":


    def test_desktop():
        print("Starting Manual Tests for desktop...")
        tests = Test_desktop()
        tests.main()

        print("Switching to unittest for filemanager...")
        print("Testing filemanager")
        unittest.main()

    test_desktop()