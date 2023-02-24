import inspect
import unittest
import traceback
import shutil, os
from pathlib import Path

from process.office import SharePoint


# [NOTE] Testing files and folders should be locally stored for these tests.
# With the exception of those for streaming (encrypted) dataframes.
SHAREPOINT = SharePoint()


class Test_SharePoint:


    PATH2PROCESSDOCFOLDER = ["WFR","Process Documents"]
    PATH2PROCESSDOC = ["WFR", "Process Documents","Workforce Reduction Demographic Analysis.docx"]
    PATH2XLSX4DATAFRAME = [
            "WFR","Workforce Reduction",".test",
            "Copy of Copy of Copy of DXC WFR Selection Template_RicardoKephart (002).xlsx"]
    PATH2ENCRYPTEDXLSX4DATAFRAME = [
            "WFR","Workforce Reduction",".test",
            "ER WFR Candidate Listing_12.07.2022.xlsx"]
    TARGET_FOLDER = Path(Path(__file__).parent, "tests", "downloadTestFolder")
    

    def clear_download_testFolder(self):
        print(inspect.stack()[0][3])
        for f in os.listdir(self.TARGET_FOLDER):
            path = Path(self.TARGET_FOLDER, f)
            if path.is_dir():
                shutil.rmtree(str(path))
            else:
                path.unlink()

    def test_doc_folder_download(self):
        self.clear_download_testFolder()
        
        print(inspect.stack()[0][3])
        SHAREPOINT.doc_folder_download(self.PATH2PROCESSDOCFOLDER, self.TARGET_FOLDER)
        
        if Path(self.TARGET_FOLDER, "Process Documents").exists():
            result = "OK"
        
        else:
            result = "FAILED!"
        
        print(f"Test_doc_folder_download: {result}")

    def test_doc_file_download(self):
        print(inspect.stack()[0][3]) 
        
        SHAREPOINT.doc_file_download(self.PATH2PROCESSDOC, self.TARGET_FOLDER)
        
        if Path(self.TARGET_FOLDER, "Process Documents", 
                "Workforce Reduction Demographic Analysis.docx").exists():
            result = "OK"
        
        else:
            result = "FAILED!"
        
        print(f"Test_docFileDownload: {result}")

    def test_doc_xlsx2dataframe(self):
        print(inspect.stack()[0][3])
        
        try:
            df = SHAREPOINT.doc_xlsx2dataframe(self.PATH2XLSX4DATAFRAME)
            
            print(df.head())
            print("OK")
        
        except Exception as e:
            traceback.print_tb(e)
            print("FAILED!")

    def test_doc_encrypted_xlsx2dataframe(self):
        print(inspect.stack()[0][3])
        
        try:
            df = SHAREPOINT.doc_encrypted_xlsx2dataframe(
                    self.PATH2ENCRYPTEDXLSX4DATAFRAME,
                    password="DXCWFR2022")
            
            size = len(df) # Empty df returned if exception thrown.
            # [NOTE] dataframes throw throw an exception if placed in a if statment.
            if size > 0: 
                print(df.head())
                print("OK")
            else:
                print("FAILED!")
        
        except Exception as e:
            traceback.print_tb(e)
            print("FAILED!")

    def test_get_relative_webpath(self):
        print(inspect.stack()[0][3])
        relwebpath = SHAREPOINT.get_relative_webpath(self.PATH2PROCESSDOCFOLDER)
        actualurl = "https://dxcportal.sharepoint.com/sites/CSSAffirmativeAction/Shared Documents/WFR/Process Documents"
        testurl = f"https://dxcportal.sharepoint.com{relwebpath}"
        print(testurl)
        result = "OK" if actualurl == testurl else "FAILED!"
        print(result)

    def main(self):
        self.test_doc_folder_download()
        self.test_doc_file_download()
        self.test_get_relative_webpath()
        self.test_doc_xlsx2dataframe()
        self.test_doc_encrypted_xlsx2dataframe()
        print("All tests have completed. Please review the print outs for results.")

test_sharepoint = Test_SharePoint()


class Test_SharePoint_Objs:  


    PATH2PROCESSDOCFOLDER = ["WFR","Process Documents"]
    PATH2PROCESSDOC = ["WFR", "Process Documents","Workforce Reduction Demographic Analysis.docx"]  


    def test_doc_file(self):
        print(inspect.stack()[0][3])
        file = SHAREPOINT.doc_file(self.PATH2PROCESSDOC)
        result = "OK" if file.name==self.PATH2PROCESSDOC[-1] else "FAILED!"
        print(result)
        
    def test_doc_folder(self):
        print(inspect.stack()[0][3])
        folder = SHAREPOINT.doc_folder(self.PATH2PROCESSDOCFOLDER)
        result = "OK" if folder.name==self.PATH2PROCESSDOCFOLDER[-1] else "FAILED!"
        print(result)
        
    def test_doc_folders(self):
        print(inspect.stack()[0][3])
        folders = SHAREPOINT.doc_folders(self.PATH2PROCESSDOCFOLDER[:-1])
        result = "OK" if len(folders) > 0 else "FAILED!"
        print(result)
    
    def test_doc_files(self):
        print(inspect.stack()[0][3])
        files = SHAREPOINT.doc_files(self.PATH2PROCESSDOCFOLDER)
        result = "OK" if len(files) > 0 else "FAILED!"
        print(result)
        
    def test_shared_documents(self):
        print(inspect.stack()[0][3])
        shareddocuments = SHAREPOINT.shared_documents()
        result = "OK" if shareddocuments.name=="Shared Documents" else "FAILED!"
        print(result)
        
    def main(self):
        self.test_doc_file()
        self.test_doc_folder()
        self.test_doc_folders()
        self.test_doc_files()
        self.test_shared_documents()


test_sharepoint_objs = Test_SharePoint_Objs()

        
if __name__ == "__main__":
    
    print("Starting Test_SharePoint")
    
    site = "CSSAffirmativeAction"
    print(f"Testing User Access to {site}")
    SHAREPOINT.site = site # printout from site assignment
    
    test_sharepoint.main()
    test_sharepoint_objs.main()