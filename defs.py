from findFixFunctions import *
from findFixHeader import *

RESULTS_DIRECTORY = "File-Crawler-Results"
FILE_CRAWLER = "File-Crawler"

# NOTE: Excel worksheet name must be <= 31 characters
LIST_ALL = "List All Files"
SPACE_FILE_FIND = "Space Error (File)"
SPACE_FOLDER_FIND = "Space (Folder)"
BAD_CHARACTER_FILE_FIND = "Bad Character Error (File)"
BAD_CHARACTER_FOLDER_FIND = "Bad Character (Folder)"
OLD_FILE_FIND = "Old File Error"
EMPTY_DIRECTORY_FIND = "Empty Directory"
EMPTY_FILE_FIND = "Empty File Error"
CHARACTER_LIMIT_FIND = "Character Limit Error"
DUPLICATE_CONTENT_FIND = "Identical Files Error"
# DUPLICATE_NAME_FIND = "Duplicate Names Error"
FILE_EXTENSION_SUMMARY = "Summarize File Types"

NULL_OPTION = "---"
SPACE_FIX = "Replace Space with Hyphen"
SEARCH_AND_REPLACE = "Search & Replace Characters~"
DELETE_OLD_FILES = "Delete Old Files~"
DELETE_EMPTY_DIRECTORIES_FIX = "Delete Empty Directories~"
DELETE_EMPTY_FILES = "Delete Empty Files"

FIND_PROCEDURES = {
    LIST_ALL: FindProcedure(LIST_ALL, listAll, True),
    SPACE_FILE_FIND: FindProcedure(SPACE_FILE_FIND, spaceFileFind, True),
    BAD_CHARACTER_FILE_FIND: FindProcedure(BAD_CHARACTER_FILE_FIND, badCharFileFind, True),
    OLD_FILE_FIND: FindProcedure(OLD_FILE_FIND, oldFileFind, True),
    EMPTY_DIRECTORY_FIND: FindProcedure(EMPTY_DIRECTORY_FIND, emptyDirectoryConcurrent, True, isFileFind=False),
    EMPTY_FILE_FIND: FindProcedure(EMPTY_FILE_FIND, emptyFileFind, True),
    SPACE_FOLDER_FIND: FindProcedure(SPACE_FOLDER_FIND, spaceFolderFind, True, isFileFind=False),
    BAD_CHARACTER_FOLDER_FIND: FindProcedure(BAD_CHARACTER_FOLDER_FIND, badCharFolderFind, True, isFileFind=False),
    CHARACTER_LIMIT_FIND: FindProcedure(CHARACTER_LIMIT_FIND, overCharLimitFind, True),
    DUPLICATE_CONTENT_FIND:FindProcedure(DUPLICATE_CONTENT_FIND, duplicateContentConcurrent, False, duplicateContentPost),
    # DUPLICATE_NAME_FIND: FindProcedure(DUPLICATE_NAME_FIND, duplicateNameConcurrent, False, duplicateNamePost),
    FILE_EXTENSION_SUMMARY: FindProcedure(FILE_EXTENSION_SUMMARY, fileExtensionConcurrent, False, fileExtensionPost)
}

FIX_PROCEDURES = {
    NULL_OPTION: None,
    SPACE_FIX: FixProcedure(SPACE_FIX, spaceFixLog, spaceFixModify, True),
    SEARCH_AND_REPLACE: FixProcedure(SEARCH_AND_REPLACE, searchAndReplaceLog, searchAndReplaceModify, True, twoStringsValidator, ">"),
    DELETE_OLD_FILES: FixProcedure(DELETE_OLD_FILES, deleteOldFilesLog, deleteOldFilesModify, True, minimumIntToInfinityOrMaxValidator, 1, "# Days Last Accessed"),
    DELETE_EMPTY_DIRECTORIES_FIX: FixProcedure(DELETE_EMPTY_DIRECTORIES_FIX, deleteEmptyDirectoriesLog, deleteEmptyDirectoriesModify, False, minimumIntToInfinityValidator, 0, "# Files Contained"),
    DELETE_EMPTY_FILES: FixProcedure(DELETE_EMPTY_FILES, deleteEmptyFilesLog, deleteEmptyFilesModify, True, columnName="Staged for Deletion")
}
