import os
import shutil
import tempfile
import pywintypes
from typing import Union
from pathlib import Path
from zipfile import ZipFile
from dataclasses import dataclass
from typing import Union, List, Optional, Tuple
from win32con import OFN_EXPLORER, OFN_ALLOWMULTISELECT
from win32gui import GetDesktopWindow, GetOpenFileNameW, GetSaveFileNameW

@dataclass # Auto init with properties as parameters, if required.
class LocalPaths:

    """
    _a dataclass class, Parses and stores local folder relative to user home._
    """

    UserDownloads: Path=Path(Path.home(), "Downloads")
    OneDriveCommercial: Path=Path(os.getenv("ONEDRIVECOMMERCIAL"))
    OneDriveConsumer: Path=Path(os.getenv("ONEDRIVECONSUMER"))

    @property
    def desktop(self):
        """_Returns the real local path for the User's Desktop._

        Returns:
            _Path_: _pathlib.Path_
        """
        return self.resolve_path(Path(Path.home(), "Desktop"))

    @property
    def documents(self):
        """_Returns the real local path for the User's Documents._

        Returns:
            _type_: _description_
        """
        return self.resolve_path(Path(Path.home(), "Documents"))


class FileManager(LocalPaths):
    """
    _A local FileManager to resolve OneDrive and OneDriveCommerical relative path issues,
     get temporary file and foler locations and for rename existing paths._
    """

    def open_file_dialog(self, title: str="DXC Office", 
                         starting_dir: Union[str, Path]=Path(Path.home(), "Documents"),
                         ext: Union[tuple, str] = "", multiselect: bool = False) -> Path:
        """_Opens a file dialog and returns the path._

        Args:
            title (str, optional): _description_. Defaults to "DXC Office".
            starting_dir (Union[str, Path], optional): _description_. Defaults to Path(Path.home(), "Documents").
            ext (Union[tuple, str], optional): _description_. Defaults to "".
            multiselect (bool, optional): _description_. Defaults to False.

        Raises:
            IOError: _as named_

        Returns:
            Path: _pathlib.Path_
        """

        if ext is None:
            ext = "All Files\0*.*\0"
        else:
            ext = "".join([f"{name}\0*.{extension}\0" for name, extension in ext])

        flags = OFN_EXPLORER
        if multiselect: flags = flags | OFN_ALLOWMULTISELECT
        
        try:
            file_path, _, _ = GetOpenFileNameW(
                                InitialDir=starting_dir,
                                Flags=flags,Title=title,
                                MaxFile=2**16,
                                Filter=ext,DefExt=ext)
            paths = file_path.split("\0")

            if len(paths) == 1:
                return paths[0]
            else:
                for i in range(1, len(paths)):
                    paths[i] = Path(paths[0], paths[i])
                paths.pop(0)

            return paths

        except pywintypes.error as e:
            if e.winerror != 0:
                raise IOError()

    def resolve_path(self, path: Union[str, Path]) -> Path:
        """_Resolves the real path, relative to the user's home folder._

        Args:
            path (Union[str, Path]): _path to resolve_

        Returns:
            Path: _office.Path_
        """
        # fixes case related errors in path names
        path = Path(str(Path(path)).lower()) 

        path = Path(path)
        if path.parent.exists():
            return path
        else:
            userHome = str(Path.home())
            parentFolder = str(path.parent)
            path2Relative2UserHome = parentFolder.split(userHome)[-1]
            directories = path2Relative2UserHome.split("\\")

            # [NOTE] Returns None if path is not resolved.
            consumerPath = Path(os.getenv("ONEDRIVECONSUMER"), (directories))
            if consumerPath.exists():
                return consumerPath
            
            commercialPath = Path(os.getenv("ONEDRIVECOMMERCIAL"), (directories))
            if commercialPath.exists():
                return commercialPath
            
    def rename_existing_path(self, path: Union[str, Path]) -> str:
        """_Returns a recursively renamed path, 
            using the _# convention, if the path 
            already exists._

        Args:
            path (str || Path): _str or pathlib.Path_

        Returns:
            pathlib.Path: _a renamed path that does not exist on the system_
        """
        path = Path(path)
        path = self.resolve_path(path)
        if path.exists():
           i = 0
           while path.exists():
               i += 1
               path = Path(path.parent, f"{path.name} ({i})")
        return path 
              
    def extract_all(self, source: Union[str, Path],
                    destination: Union[str, Path]=None):
        """_Extracts the contents of a zip folder_

        Args:
            source (_type_): _description_
            destination (_type_): _description_
        """
    
        source = self.resolve_path(source)
        if not destination:
            destination = source.parent

        if not destination.exists():
            destination.mkdir(parents=True)
        
        with ZipFile(str(source), 'r') as ref:
            ref.extractall(str(destination))

    def remove_directory(self, path: Union[str, Path]) -> bool:
        """_Removes a populated directory. Warning, deleted folders
            will not be found in the Recycling bin. The removal
            of directories is permanent._

        Args:
            path (Union[str, Path]): _C:\\path\to\directory_

        Returns:
            bool: _whether the operation was successful._
        """
        path = self.resolve_path(path)
        
        if path.exists():
            shutil.rmtree(str(path))
        else:
            print("Nothing done. Path does not exist.")

        return not path.exists()

    def get_temp_file(self, suffix: str):
        """
        _Unlike tempfile.TemporaryFile, NamedTemporaryFile
         which this function wraps, returns a file that is
         guaraurteed to have a visible name in the file system._

        Args:
            suffix (str): _file.suffix_

        Returns:
            _tempfile._TemporaryFileWrapper_: _tempfile._TemporaryFileWrapper_
        """
        
        return tempfile.NamedTemporaryFile(suffix=suffix)

    def get_secured_temp_file(self, suffix: str):
        """
        _Returns a secured file. Here secured means that it 
         can only be used by the creating UserID._

        Args:
            suffix (str): _file.suffix_

        Returns:
            _str_: _secured file location_
        """

        return tempfile.mkstemp(suffix=suffix)

    def get_in_memory_temp_file(self, suffix: str, bytes: bool = True):
        """
        _Returns a temp file in memory. In memory files are written to
         disk when memory is exceeded._

        Args:
            suffix (str): _file.suffix_
            bytes (bool, optional): _description_. Defaults to True.

        Returns:
            _tempfile.SpooledTemporaryFile_: _tempfile.SpooledTemporaryFile_
        """

        mode = "w+b" if bytes else "w+"
        return tempfile.SpooledTemporaryFile(suffix=suffix, mode=mode)
    
    def get_temp_Folder(self, ignore_cleanup_errors: bool=True):
        """
        _Returns a temp folder just like mkdtemp._

        Args:
            ignore_cleanup_errors (bool, optional): _description_. Defaults to True.

        Returns:
            _type_: _description_
        """

        return tempfile.TemporaryDirectory(ignore_cleanup_errors=ignore_cleanup_errors)
