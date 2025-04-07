import os
from workbookManager import WorkbookManager
import string
from datetime import datetime
import hashlib
from getFileOwner import getOwnerCatch



# Used by badCharFind
# Includes ' ' (space) as there is a separate procedure for finding that error
PERMISSIBLE_CHARACTERS = set(string.ascii_letters + string.digits + "- ")
CHARACTER_LIMIT = 200

# Used by oldFileFind and deleteOldFiles
TODAY = datetime.now()



def setWorkbookManager(newManager: WorkbookManager):
    # Globally declare the WorkbookManager object
    global wbm
    wbm = newManager



def addLongPathPrefix(dirAbsolute):
    if dirAbsolute.startswith('\\\\'):
        return '\\\\?\\UNC' + dirAbsolute[1:]
    else:
        return '\\\\?\\' + dirAbsolute



def writeOwnerHeader(ws):
    # importGetOwner()
    ws.write(0, wbm.AUXILIARY_COL, "Owner", wbm.headerFormat)
    


def joinDirToFileName(dirAbsolute, fileName):
    return dirAbsolute + "\\" + fileName