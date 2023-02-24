import os
import traceback
from time import sleep
from io import BytesIO
from pathlib import Path
from typing import Union
from html import unescape
from getpass import getpass

import msoffcrypto
import pandas as pd
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint import folders, files
from office365.sharepoint.folders import folder
from office365.sharepoint.files import file


"""[o365 Summary] 
    A minimal wrapper around pyOffice365RestApiClient.
    Created with ease of use and interoperability in mind.

    When making changes to this file, please note
    that the words file, files, folder, folders are reserved
    words. Any variables referring to an instance there of
    should be renamed or prefixed with a 'x'.
"""


class AuthenticationFailed(Exception):
    pass


class SiteNotAssigned(Exception):
    pass


class FailedToFindParentFolder(Exception):
    pass


class FileNotFoundInParentDirectory(Exception):
    pass


class FolderNotFoundInParentDirectory(Exception):
    pass


class FolderEmptyOfContent(Exception):
    pass


class DownloadFailed(Exception):
    pass


class PathNotFoundError(Exception):
    pass


class Api:
    
    
    """
    _Parent class for o365 sharepoint and future api classes.
     Authentication, reauthentication, and parsing of parameters
     & values into human readable format, done from here.
     
     User signin is done at the tenant level, after which point
     site can be reassigned with auto cascading changes to the 
     web context using credentials stored in the form a 
     office365.runtime.auth.user_credential.UserCredential.
     The username is stored as a class property, the password
     is only used to generate the UserCredential obj and
     is not stored.
     
     Several essential functions related to local file management
     are included in the Api rather than as imported functions of
     office.FileManager to avoid cirular import. The same 
     functionailty can be found in office.FileManager and should
     be imported from there for use outside of one of the o365
     api classes._
    """
    
    
    def __init__(self, username: str = None, password: str = None):     
        self.get_credentials((username,password))
        del username, password
        self.site = None # simulates tenant login.

    def __load_ctx(self):
        """_sets .ctx and verifies connection_

        Raises:
            AuthenticationFailed: _failed to get ctx_
        """
        try:
            url = f"https://COMPANYPORTAL.sharepoint.com/" + \
                    f"sites/{self.site}"
            credentials = self.credentials
            setattr(
                self, "ctx", 
                ClientContext(url).with_credentials(credentials))
            web = self.ctx.web
            self.ctx.load(web).execute_query()
            print(f"Now connected to: {web.title}")
        except Exception as e:
            traceback.print_exception(e)
            raise AuthenticationFailed
    
    def __get_relative_webpath_from_url(self, url: str) -> str:
        """_Parses the relative web path._

        Args:
            url (_str_): _https://web/path_

        Returns:
            _str_: _relative web path_
        """
        url = url.split("https://")[-1]
        url = unescape(url)
        urlParts = url.split("/")
        self.tenant = urlParts[0].split(".")[0]
        self.site = urlParts[2]
        path = f"/sites/{self.site}"
        try:
            for xfolder in urlParts[3:]:
                path += f"/{xfolder}"
            return path
        except:
            pass
    
    def __get_relative_webpath_from_pathlist(self, directoryList: list) -> str:
        """_Parses the relative web path._

        Args:
            directoryList (_list_): _list of directories_

        Returns:
            _str_: _relative web path_
        """
        path = f"/sites/{self.site}/Shared Documents"
        for xfolder in directoryList:
            path += f"/{xfolder}"
        return path
    
    @property
    def site(self):
        return self._site    
    
    @site.setter
    def site(self, value):
        """_Assignment of .site -> .loadCtx._

        Args:
            value (_str_): _name of new site_
        """
        self._site = value
        if value:
            print(f".site changed to {value}")
        if self._site:
            print("Loading Web Context...")
            self.__load_ctx()
            
    def get_credentials(self, auth: tuple=(None, None)):
        """_Resets UserCredentials._

        Args:
            auth (_tuple_): _(username,password)_
        """
        setattr(self,"tenant","dxcportal")
        username = auth[0] if auth[0] else input("Global Source Username: ").strip()
        if not "@" in username:
                username = f"{username}@dxc.com"
        if auth[1]:
            setattr(self,"credentials",
                    UserCredential(username, auth[1])) 
        else:
            setattr(self,"credentials",
                    UserCredential(
                        username, getpass("Global Source Password: ").strip()))            
        print(f"UserCredentials has been assigned for {username}.")
        del auth, username
            
    def get_relative_webpath(self, urlPathlist) -> str:
        """_Combines .__get_relative_webpath_from_url & .__get_relative_webpath_from_pathlist._

        Args:
            urlPathlist (_str | list_): _url | a list of directories_

        Returns:
            _str_: _relative web path_
        """
        if type(urlPathlist) == str and "/" in urlPathlist:
            return self.__get_relative_webpath_from_url()
        elif type(urlPathlist) == str:
            urlPathlist = [urlPathlist]
        return self.__get_relative_webpath_from_pathlist(urlPathlist)
            
    def reload_ctx(self):
        """_Reload of web context. Useful for network disruptions._

        Args:
            tenant (str, optional): _None will reload with existing .domain.TENANT_. Defaults to None.
            site (str, optional): _None will reload with existing .site_. Defaults to None.
        """
        self.__load_ctx()


class SharePoint(Api):

    """
    _A easy to use wrapper around the SharePoint api. Single function downloads. 
     Returns file and folder api objects that have been pre initialized
     and fetched. The relative web location for requested assets is resolved via 
     url or a human-readable Pathlist, [a, list, of, directories, from, Shared 
     Documents, tofile.xlsx]. doc_folder_download recreates the file and folder 
     structure for the requested asset. There is a single point of download 
     for both file and folder download functions with a named exception 
     class for failed download requests.

     From within class Process(Process):

     '''python
     # Download via url, default downloads to user's Downloads folder
     url = "https://dxcportal.sharepoint.com/:u:/r/sites/x/Shared%20Documents/Assets/x.pub"
     download_path = self.sharepoint.doc_file_download(url)
     print(str(download_path))
     print(download_path.exists())

     # Download via list relative to Shared Documents, local path to destkop is autoresolved if nested in OneDrive
     desktop = Path(Path.home(), "Desktop")
     download_path = sp.doc_folder_download(["Assets"],desktop)
     # download_path: %OneDriveCommercial%\Desktop
     ```

     SharePoint file, files, folder, and folders objects cannot be returned outside
     of their class.method context or else an Exception will be thrown. To avoid this
     always be sure to return .get().execute_query() directly:

     ```python
     def get_files_from_folder(self, urlPathlist)
         xfolder = self.sharepoint.doc_folder()
         xfiles = xfolder.files
         return xfiles.get().execute_query()

     def print_folder_file_names(self)
         urlPathlist=['folders','relative','shared','documents']
         fetched_files = self.get_files_from_folder(urlPathlist)
         for file in fetched_files:
             print(file.name)
     ```

     The class can be initialized with or without a username/password. If either 
     is not provided, both will be requested at init from the command line. Although
     credentials are fetched and stored at init, authentication happens on a site by
     site basis with the assignment of self.site, provided the user has access.
     
     Local paths for downloads are always auto resolved prior to initiating the download
     request, whether the directory is nested in a relative OneDrive folder or the 
     download target already exists. Post download, the resolved path is returned._
    """

    def __init__(self, username: str = None, password: str = None):
        """
        Args:
            username (str, optional): _username@dxc.com_. Defaults to None.
            password (str, optional): _globalpass123!_. Defaults to None.
        """

        super().__init__(username, password)

    def __create_local_download_folder(self, folderName: str, parentFolder: Union[str, Path]=None) -> Path:
        """_Hidden, creates a local folder for download to avoid downstream conflicts._

        Args:
            folderName (str): _requested name for the folder, this name may be appended, if existing_
            localPath (Union[str, Path], optional): _the. Defaults to None.

        Returns:
            Path: _description_
        """

        if parentFolder:
            localPath = Path(parentFolder, folderName)
            localPath = self.__resolve_onedrive_path(localPath)
            localPath = self.__resolve_existing_path(localPath)
            
            print(f"SharePoint: Creating {str(localPath)}")
            localPath.mkdir()
            return localPath
        else:
            return Path(Path.home(), "Downloads")

    def __resolve_onedrive_path(self, path: Union[str, Path]):
        """_Hidden, gets the real location of a foler relative to user's home folder and OneDrive._

        Args:
            path (Path): _a folder or file nested within folder which usually reside in 
                          the %USERPROFILE% folder_

        Raises:
            PathNotFoundError: _could not find a real path for the requested path._

        Returns:
            _type_: _description_
        """

        # [NOTE] For folders inside folders that usually reside inside the user's
        # home directory, i.e. Desktop.
        path = Path(str(path).lower()) # resolves case related issue with filenames on windows
        if path.parent.exists():
            return path
        
        else:
            directories = str(path.parent).split(str(Path.home()))[-1].split("\\")

            consumerOneDrivePath = Path(os.getenv("ONEDRIVECONSUMER"), (directories))
            if consumerOneDrivePath.exists(): return consumerOneDrivePath
            
            commercialOneDrivePath = Path(os.getenv("ONEDRIVECOMMERCIAL"), (directories))
            if commercialOneDrivePath.exists(): return commercialOneDrivePath

            raise PathNotFoundError
                
    def __resolve_existing_path(self, path: Path, i: int=0) -> str:
        """_Returns a recursively renamed path, using the _# convention,
            if the path already exists._

        Args:
            path (str || Path): _str or pathlib.Path_

        Returns:
            pathlib.Path: _a renamed path that does not exist on the system_
        """

        path = self.__resolve_onedrive_path(path)
        if path.exists():
            if "_" in path.name:
                
                try:
                    i = int(path.name.split("_")[-1])
                    path = Path(
                        path.parent, 
                        "".join(path.name.split(f"_{i}")[:-1])+f"_{i+1}")
                    return self.rename_existing_path(path)
                
                except:
                    return Path(path.parent, f"{path.name}_1")
        
        else:
            return path

    def __doc__2dataframe(self, urlPathlist: Union[str, list],
                          sheet_name: str=None, sheet_number: int=None) -> pd.DataFrame:
        # Streams a Excel file stored in SharePoint directly into a pandas.DataFrame.
        # Args:
        #    urlPathlist (Union[str, list]): _description_
        #    sheet_name (str, optional): _description_. Defaults to None.
        #    sheet_number (int, optional): _description_. Defaults to None.
        # Returns:
        #    pd.DataFrame: _description
        xfile, iobuff = self.doc_file(urlPathlist), BytesIO()
        xfile.download(iobuff).execute_query()
        
        if sheet_number:
            sheet_name = sheet_number # doesn't take into account conflicts    
        
        else:
            sheet_name = 0
        
        try:
            return pd.read_excel(iobuff, sheet_name=sheet_name)
        
        except:

            try:
                return pd.read_csv(iobuff)
            
            except:
                print("File Types: csv, xls, xlsx, xlsm, xlsb, odf, ods and odt")
    
    def doc_folder_upload(self, urlPathlist: Union[str, list], localPath: Union[str, Path]):
        """_summary_

        Args:
            urlPathlist (Union[str, list]): _description_
            localPath (Union[str, Path]): _description_
        """
        
        relWebPath = ""
        if type(urlPathlist) == str:
            relWebPath = f"Shared Documents\{urlPathlist.split('Shared%Documents')[-1]}"

        else:
            for dir in urlPathlist: relWebPath += f"\{dir}"

        print("Creating the Directory Structure...")
        for (xfolder, _, _) in os.walk(localPath):
            nestedWebPath = f"{relWebPath}\{xfolder}"
            
            xfolder = self.web.ensure_folder_path(nestedWebPath).execute_query()
            print(f"{xfolder.serverRelativeUrl}\033[0K\r") # serves as a progress meter

        print("Uploading Files...")
        for (xfolder, folders, xfiles) in os.walk(localPath):
            
            for xfile in xfiles:
                targetList = f"{relWebPath}\{xfolder}\{xfile}".split("\\")
                localTarget = Path(localPath, (folders), xfile) 
                self.doc_file_upload(targetList, localTarget)

    
    def doc_file_upload(self, pathlist: list, localPath: Union[str, Path]):
        """_Upload a file to the requested target._

        Args:
            pathlist _list_ : _list of nested directories_
            localPath _Union[str, Path]_ : _a local file path_

        Raises:
            FileNotFoundError: _as named_
        """
        
        localPath = Path(localPath)
        if not localPath.exists():
            raise FileNotFoundError

        with open(localPath, "rb") as f:
            contents = f.read()
            xfolder = self.doc_folder(pathlist)

            xfile = xfolder.upload_file(localPath.name, contents).execute_query()
            xfile.get().execute_query()
            
            # serves as a progress meter for doc_folder_upload
            print(f"{xfile.name}\033[0K\r")

    def doc_share(self, urlPathlist: Union[str, list], share_type: str="OrganizationView"):
        """_Fetchs the share url for a document file, in the share_type requested.
        
            share_types: {
                'Uninitialized'    : 0,
                'Direct'           : 1,
                'OrganizationView' : 2,
                'OrganizationEdit' : 3,
                'AnonymousView'    : 4,
                'AnonymousEdit'    : 5,
                'Flexible'         : 6}
         
            share_type_descriptions: {
                0  : "A value has not been initialized",
                1  : "A direct link or canonical URL to an object",
                2  : "An organization access link with view permissions to an object",
                3  : "An organization access link with edit permissions to an object",
                4  : "An anonymous access link with view permissions to an object",
                5  : "An anonymous access link with edit permissions to an object",
                6  : "A tokenized sharing link where properties can change without affecting link URL"}_

        Args:
            urlPathlist (Union[str, list]): _description_
            share_type (str, optional): _description_. Defaults to "OrganizationView".

        Returns:
            _type_: _description_
        """

        try:
            asset = self.doc_file(urlPathlist)
        
        except:
            try:
                asset = self.doc_folder(urlPathlist)
                result = asset.share_link(self.SHARE_TYPES[share_type]).execute_query()
                info = asset.get_sharing_information().execute_query()
        
                print("Share Url Properties:", info.properties)
                return result.value.sharingLinkInfo.Url
            
            except Exception as e:
                traceback.print_exception(e)
                print("Are you requesting the url for a file?")
        
    def doc_unshare(self, urlPathlist: Union[str, list], share_type: str="all"):
        """_Unshares url for a document file, in the share_type requested.
        
            share_types: {
                'Uninitialized'    : 0,
                'Direct'           : 1,
                'OrganizationView' : 2,
                'OrganizationEdit' : 3,
                'AnonymousView'    : 4,
                'AnonymousEdit'    : 5,
                'Flexible'         : 6,
                'all': loops through all the above share_types}
         
            share_type_descriptions: {
                0  : "A value has not been initialized",
                1  : "A direct link or canonical URL to an object",
                2  : "An organization access link with view permissions to an object",
                3  : "An organization access link with edit permissions to an object",
                4  : "An anonymous access link with view permissions to an object",
                5  : "An anonymous access link with edit permissions to an object",
                6  : "A tokenized sharing link where properties can change without affecting link URL"}_

        Args:
            urlPathlist (Union[str, list]): _description_
            share_type (str, optional): _description_. Defaults to "OrganizationView".

        Returns:
            _type_: _description_
        """

        try:
            asset = self.doc_file(urlPathlist)
        
        except:
            try:
                asset = self.doc_folder(urlPathlist)
            
            except Exception as e:
                traceback.print_exception(e)
                
        if share_type == "all":
            for opt in self.SHARE_TYPES.values():
                
                try:
                    asset.unshare_link(opt).execute_query()
                except:
                    pass
        else:
            asset.unshare_link(self.SHARE_TYPES[share_type])

  
    def doc_folder_download(self, urlPathlist: Union[str,list],
                          parentFolder: Union[str, Path]=Path(Path.home(), "Downloads"),
                          _0: Path=None):
        """_Downloads a sharepoint folder and all of its contents to a specified parent folder_

        Args:
            urlPathlist (Union[str,list]): _description_
            localPath (Union[str, Path]): _description_
            _0 (Path, optional): _description_. Defaults to None.

        Returns:
            _type_: _description_
        """

        # [NOTE] Destination path modification for folders moved into onedrive.
        # The parent must exist for the download.
        relWebPath = self.get_relative_webpath(urlPathlist)
        xfolder = self.ctx.web.get_folder_by_server_relative_path(relWebPath)
        xfolder.get().execute_query()
        
        files = xfolder.files
        files.get().execute_query()
        
        name = xfolder.name
        localPath = self.__create_local_download_folder(name, parentFolder)
        # [TODO] change algo using pointers return the starting point.
        # [PATCH] for return of original download folder path
        if not _0: _0 = localPath 

        for f in files:
            try:
                filename = f.name
                # [NOTE] Assignment to _ required or the downloaded path is printed.
                _ = self.doc_file_download(urlPathlist+[filename], Path(localPath))
            except:
                print(f"Failed to download: {f.name}")
        
        xfolders = xfolder.folders
        xfolders.get().execute_query()
        for f in xfolders:
            subFolder = urlPathlist + [f.name]
            self.doc_folder_download(subFolder, localPath, _0=_0)
            if f == xfolders[-1]:
                return _0

    def doc_file_download(self, urlPathlist: Union[str, Path],
                        localDir: Union[str, Path]=Path(Path.home(), "Downloads")):
        """_Downloads a document file._

        Args:
            urlPathlist (Union[str, Path]): _a url string or list of directories to file_
            localPath (Union[str, Path], optional): _local directory for download_. Defaults to None.

        Raises:
            FailedToFindParentFolder: _BUG: local path resolution issue_
            DownloadFailed: _Network or Api issue_
        """

        relWebPath = self.get_relative_webpath(urlPathlist)
        xfile = self.ctx.web.get_file_by_server_relative_path(relWebPath)
        xfile.get().execute_query()
        
        localPath = Path(localDir, xfile.name)
        # [NOTE] Of the directories in the User's Home folder, Downloads is the only one
        # that remains in it's original location upon turning on OneDrive. Downloads paths 
        # pointing to Desktop and Documents must first be resolved as they reside in either 
        # OneDriveConsumer or OneDriveCommercial.
        
        # [__resolve_onedrive_path] Destination path modifications for folders moved into onedrive:
        # The parent must exist for the download, else __resolve_onedrive_path returns
        # None.
        localPath = self.__resolve_onedrive_path(localPath)
        if not localPath:
            raise FailedToFindParentFolder
        
        # [__resolve_existing_path] In cases where the download destination already exists, 
        # a incrementing number is appended to the localPath to make downloading possible.
        localPath = self.__resolve_existing_path(localPath)
        
        # The only download point using context manager.
        with open(localPath, "wb+") as f:
            # Self erasing updates also serve as a unobstrusive 
            # progress monitor for long downloads.
            print(f"Downloading: {str(localPath)}\033[0K\r")
            try:
                xfile.download(f).execute_query()
                print(f"Completed.\033[0K\r")
            except Exception as e:
                traceback.print_exception(e)
                raise DownloadFailed
        return localPath
                    
    def shared_documents(self):
        """_returns the Shared Documents folder_

        Returns:
            _folder_: _office365.sharepoint.folders.folder_
        """

        try:
            shared_documents = self.ctx.web.default_document_library().root_folder
            return shared_documents.get().execute_query()
        
        except Exception as e:
            traceback.print_exception(e)

    def doc_folder(self, urlPathlist) -> folder:
        """
        _returns a office365.sharepoint.folder obj_

        Args:
            urlPathlist (_str | list_): _url | a list of directories_

        Returns:
            _folder_: _office365.sharepoint.folders.folder_
        """

        try:
            relPath = self.get_relative_webpath(urlPathlist)
            xfolder = self.ctx.web.get_folder_by_server_relative_path(relPath)
            return xfolder.get().execute_query()
        except Exception as e:
            traceback.print_exception(e)

    def doc_folders(self, urlPathlist) -> folders:
        """
        _returns a office365.sharepoint.folders obj_

        Args:
            urlPathlist (_str | list_): _url | a list of directories_

        Returns:
            _folders_: _office365.sharepoint.folders_
        """

        relPath = self.get_relative_webpath(urlPathlist)
        mainfolder = self.ctx.web.get_folder_by_server_relative_path(relPath)
        mainfolder.get().execute_query()

        xfolders = mainfolder.folders
        return xfolders.get().execute_query()
            
    def doc_files(self, urlPathlist) -> files:
        """
        _returns a files office365.sharepoint.files obj_

        Args:
            urlPathlist (_str | list_): _url | a list of directories_

        Returns:
            _files_: _office365.sharepoint.files_
        """

        relPath = self.get_relative_webpath(urlPathlist)
        mainfolder = self.ctx.web.get_folder_by_server_relative_path(relPath)
        mainfolder.get().execute_query()

        xfiles = mainfolder.files
        return xfiles.get().execute_query()
    
    def doc_file(self, urlPathlist) -> file:
        """_summary_

        Args:
            urlPathlist (_str | list_): _url | a list of directories_

        Returns:
            _file_: _office365.sharepoint.files.file_
        """

        relPath = self.get_relative_webpath(urlPathlist)
        xfile = self.ctx.web.get_file_by_server_relative_path(relPath)
        return xfile.get().execute_query()

    def doc_csv2dataframe(self, urlPathlist: Union[str, list]) -> pd.DataFrame:
        return self.__doc__2dataframe(self, urlPathlist)

    def doc_xlsx2dataframe(self, urlPathlist: Union[str, list]) -> pd.DataFrame:
        return self.__doc__2dataframe(self, urlPathlist)

    def doc_encrypted_xlsx2dataframe(self, urlPathlist: Union[str, list], password: str) -> pd.DataFrame:
        """_Streams a encrypted Excel file stored in SharePoint directly into a pandas.DataFrame.
        
            [NOTE] Encrypted xls_ streaming (not saving the file to disk) is restricted to the first
            Worksheet of the Workbook. This appears to be a limitation of the msoffcrypto package,
            when chaining the data into a pandas.DataFrame. There is no patch currently scheduled.

            [TODO - maybe remove] Encrypted files are temporarily placed into user %TEMP% where and
            read into memory. The file is immediately deleted upon being read. Decryption
            and loading into dataframe happens using the data stored in memory._
            
        
        Args:
            urlPathlist (Union[str, list]): _description_
            password (str): _description_
            sheet_name (str, optional): _description_. Defaults to None.
            sheet_number (int, optional): _description_. Defaults to None.

        Returns:
            pd.DataFrame: _description_
        """    
        
        xfile, iobuff = self.doc_file(urlPathlist), BytesIO()
        iobuff = BytesIO() # stores decrypted contents
        
        path = Path(os.environ["TEMP"], "temp.xlsx")
        with open(path, "wb+") as f:
            try:
                xfile.download(f).execute_query()
                decrypter = msoffcrypto.OfficeFile(f)
            except Exception as e:
                traceback.print_exception(e)
                print("Failed to decrypt file.")
        
        try:
            path.unlink()
        except:
            try:
                os.remove(path)
            except:
                msg = f"Possible Orphaned File. Location: {str(path)}. "
                msg += "Please verify that it has been deleted."

        try:
            decrypter.load_key(password=password)
            decrypter.decrypt(iobuff)
            return pd.read_excel(iobuff)

        except Exception as e:
            traceback.print_exception(e)
            print("Failed to read decrypted contents.")
            return pd.DataFrame()