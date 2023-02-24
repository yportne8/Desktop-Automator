from pathlib import Path
from typing import Union
from win32com import client


class FailedToDecrypt(Exception):
    pass


class FailedToOpen(Exception):
    pass


class OpenedDocumentNotFound(Exception):
    pass


class Word:
 
    """
    _A light wrapper around Word using win32com._
    """
    
    def __init__(self):
        # [NOTE] Setting to Word defaults.
        self._font = "Calibri"
        self._size = 11
        self._bold = False
        self._italics = False
        self._underline = False

    @property
    def font(self):
        return self._font
    
    @font.setter
    def font(self, value):
        try:
            self.app.Selection.Font.Name = value
            self._font = value
        except:
            print(f"Unknown font: {value}")

    @property
    def size(self):
        return self._size
    
    @size.setter
    def size(self, value: int):
        self._size = abs(value)

    @property
    def bold(self):
        return self._bold
    
    @bold.setter
    def bold(self, value: bool):
        self._bold = value

    @property
    def italics(self):
        return self._italics
    
    @bold.setter
    def italics(self, value: bool):
        self._italics = value

    @property
    def underline(self):
        return self._underline
    
    @bold.setter
    def underline(self, value: bool):
        self._underline = value

    def __del__(self):
        # [NOTE] Save request for opened document changes from Word.
        self.app.Application.Quit()
        print("Word has closed.")
    
    def __set_font(self):
        """_Sets the cursor font._

        Args:
            text (_type_): _The text segment to write._
            font (_type_): _Font for the text_
            size (_type_): __
            bold (bool, optional): _description_. Defaults to False.
            italic (bool, optional): _description_. Defaults to False.
            underline (bool, optional): _description_. Defaults to False.
        """

        self.app.Selection.Font.Name = self.font
        self.app.Selection.Font.Size = self.size
        self.app.Selection.Font.Bold = self.bold
        self.app.Selection.Font.Italic = self.italics
        self.app.Selection.Font.Underline = self.underline

    def __paste_into_document(self):
        """
        _Paste into document from clipboard_
        """
        cursor = self.app.Selection
        cursor.Paste()
        
    @property
    def app(self):
        # [NOTE] application is initialized and hidden when first dispatched,
        # until a document is opened, at which point it is unhidden.
        # Calling dispatch again on the opened application has no impact and
        # resolves the issue of the application being closed by ex-programmtic
        # sources.
        return client.Dispatch('Word.Application')
    
    @property
    def documentNames(self):
        return [d for d in self.app.Documents]

    @property
    def documents(self):
        return [self.app.Documents[name] for name in self.documentNames]

    @property
    def activeDocumentName(self):
        return self.app.Application.ActiveDocument()

    @property
    def activeDocument(self):
        return self.app.Documents[self.activeDocumentName]
    
    @property
    def document(self):
        return self._document
    
    @document.setter
    def document(self, value):
        self._document = value

    def get_document(self, opened_document_name: str):
        """
        _Returns the opened document object based on the 
         opened document's name_

        Args:
            opened_document_name (str): _as named_

        Raises:
            OpenedDocumentNotFound: _as named_

        Returns:
            _win32comobj_: _Word.Documents.Document_
        """

        opened_document_name = opened_document_name.split(".docx")[0]
        try:
            return self.app.Documents[opened_document_name]
        
        except:
            print(f"{opened_document_name} is closed.")
            raise OpenedDocumentNotFound
        
    def hide(self):
        """
        _Hides the window, does not close it. New windows cannot be 
         opened without unhiding all other windows._
        """

        app = self.app
        app.Visible = False

    def unhide(self):
        """
        _Unhides the window._
        """

        app = self.app
        app.Visible = True
        
    def new_document(self):
        """
        _Returns the document object for a new document._

        Returns:
            _object_: _client.Dispatch('Word.Application').Document_
        """

        return self.app.Documents.Add()

    def open(self, path: Union[str, Path], readonly: bool=False) -> object:
        """
        _Returns the document object for the opened document_

        Args:
            path (Union[str, Path]): _path to document_
            readonly (bool): _open in readonly_

        Raises:
            FileNotFoundError: _path does not exist_

        Returns:
            _object_: _client.Dispatch('Word.Application').Document_
        """

        path = Path(path)

        if not path.exists():
            raise FileNotFoundError
        
        ConfirmConversions=False
        # returns a document object
        return self.app.Documents.Open(str(path), ConfirmConversions, readonly)
    
    def write2file(self, text: str, path: Union[str, Path], keep_open: bool=False):
        """
        _Writes text into an unopened document. Procedure is done on a hidden document.
         If keep_open is set to True, then the function will return the document object,
         else save and close the document._

        Args:
            text (str): _Text to write_
            path (Union[str, Path]): _Path to .docx_
            keep_open (bool, optional): _whether to keep the document open_. Defaults to False.

        Returns:
            _win32comobj_: _Word.Documents.Document_
        """

        if not self.app.Visible:
            document = self.open(path)
            self.hide()
        else:
            document = self.open(path)
        
        self.__set_font() # sets font,
        try:
            document.Content.InsertAfter(text)
            document.Save()
            if keep_open:
                return document
            else:
                document.Close()
        except:
            msg = "Word is currently running a process is not accepting commands. A dialogue window may be open."
            print(msg)
            self.unhide()

    def write2document(self, text: str, document):
        """
        _Writes text into a opened document._

        Args:
            text (str): _Text to write_
            document (_str | win32comobj_): _document name or document obj_
        """
        self.__write_text(text)
        try:
            if type(document) == str:
                self.app.Documents[document].Select()
            else:
                document.Select()
            document.Save()
        except:
            msg = "Word is currently running a process is not accepting commands. A dialogue window may be open."
            print(msg)
            self.unhide()
            
    def paste2document(self, document):
        """
        _Pastes data copied via Word or other desktop office applications._

        Args:
            document (_str | win32comobj_): _document name or document obj_
        """

        if type(document) == str:
            document = document.split(".docx")[0]
            document = self.app.Documents[document]
        
        try:
            document.Select()
            self.__paste_into_document()
        except:
            msg = "Word is currently running a process is not accepting commands. A dialogue window may be open."
            print(msg)
            self.unhide()