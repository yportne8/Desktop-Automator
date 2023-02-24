import inspect
from .word import Word
from .excel import Excel
from .access import Access
from .outlook import Outlook
from .shared import (Console, 
    Table,FileManager,Notify,
    WindowManager,Browser)


def GetMembers(obj): return [m[0] for m in inspect.getmembers(obj) if m[0][0] != "_"]