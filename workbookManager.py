import xlsxwriter
from typing import List
from time import time, sleep
import os
import stat
import filesScannedSharedVar
from ExcelWritePackage import ExcelWritePackage
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock


class WorkbookManager:

    def __init__(self, workbookPathName: str):
        self.workbookPathName = workbookPathName
        self.wb = xlsxwriter.Workbook(workbookPathName)

        self.summarySheet = self.wb.add_worksheet("Summary")
        self.summarySheet.activate()  # view this worksheet on startup

        self.excludedDirs = []  # set within initiateCrawl()

        self.findSheets = {} # procedureObject : worksheet
        self.fixSheets = {}
        self.fixProcedureArgs = {}
        self.fixProcedureFunctions = {}
        # For summarySheet, first # rows are used for mainstay metrics. Skip a line, then write variable number of procedure metrics.
        self.sheetRows = {self.summarySheet: 14} # worksheet : Integer
        self.summaryCounts = {}  # worksheet : Integer

        # Lists to avoid many redundant if statements
        self.findProceduresConcurrentOnly = []
        self.fileFindProcedures = []
        self.fileFixProcedures = []
        self.folderFindProcedures = []
        self.folderFixProcedures = []

        # Summary metrics
        self.filesScannedCount = 0
        self.foldersScannedCount = 0
        self.fileErrorCount = 0
        self.folderErrorCount = 0
        self.executionTime = 0

        # Constant columns
        self.DIR_COL = 0
        self.ITEM_COL = 1
        self.OUTCOME_COL = 2  # describes either an Error or a Modification, depending on the procedure type
        self.AUXILIARY_COL = 3

        # Default cell styles
        self.dirFormat = self.wb.add_format({"bg_color": "#99CCFF", "bold": True})  # blueish
        self.errorFormat = self.wb.add_format({"bold": True})  # "bg_color": "#FF4444", # reddish
        self.modifyFormat = self.wb.add_format({"bg_color": "#00FF80", "bold": True})  # greenish
        self.logFormat = self.wb.add_format({"bg_color": "#9999FF", "bold": True})  # purplish
        self.headerFormat = self.wb.add_format({"bg_color": "#C0C0C0", "bold": True})  # grayish
        self.summaryValueFormat = self.wb.add_format({})

        self.warningWeakFormat = self.wb.add_format({"bg_color": "#FFEB9C", "bold": True})  # yellowish
        self.warningStrongFormat = self.wb.add_format({"bg_color": "#FFC7CE", "bold": True})  # reddish


    def getAllProcedureSheets(self):
        return (list(self.findSheets.values()) + list(self.fixSheets.values()))


    def addFindProcedure(self, findProcedureObject):
        tmpWsVar = self.wb.add_worksheet(findProcedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.getAllProcedureSheets()), 0, findProcedureObject.name + " count", self.headerFormat)
        
        self.findSheets[findProcedureObject] = tmpWsVar
        self.sheetRows[tmpWsVar] = 0  # NOTE: CHANGED HERE FROM 0 SO SHEETROW TRACKS LAST ROW WRITTEN TO
        self.summaryCounts[tmpWsVar] = 0

        if findProcedureObject.isConcurrentOnly:
            tmpWsVar.freeze_panes(1, 0)
            tmpWsVar.write(0, self.DIR_COL, "Directory", self.headerFormat)
            tmpWsVar.write(0, self.ITEM_COL, "Item", self.headerFormat)
            tmpWsVar.write(0, self.OUTCOME_COL, "Error", self.headerFormat)

            self.findProceduresConcurrentOnly.append(findProcedureObject)
        
        if findProcedureObject.isFileFind:
            self.fileFindProcedures.append(findProcedureObject)
        else:
            self.folderFindProcedures.append(findProcedureObject)
            

    def addFixProcedure(self, fixProcedureObject, allowModify, addRecommendations, arg) -> bool:
        tmpWsVar = self.wb.add_worksheet(fixProcedureObject.name)
        self.summarySheet.write(self.sheetRows[self.summarySheet] +len(self.getAllProcedureSheets()), 0, fixProcedureObject.name + " count", self.headerFormat)

        self.fixSheets[fixProcedureObject] = tmpWsVar
        self.sheetRows[tmpWsVar] = 0
        self.summaryCounts[tmpWsVar] = 0

        if allowModify:
            self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.modifyFunction
        elif addRecommendations and fixProcedureObject.recommendLogFunction:
            self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.recommendLogFunction
        else:
            self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.logFunction

        ###
        #if fixProcedureObject.recommendLogFunction:
        #    self.fixProcedureFunctions[fixProcedureObject] = fixProcedureObject.recommendLogFunction
        ###

        if fixProcedureObject.isFileFix:
            self.fileFixProcedures.append(fixProcedureObject)
        else:
            self.folderFixProcedures.append(fixProcedureObject)

        tmpWsVar.freeze_panes(1, 0)
        tmpWsVar.write(0, self.DIR_COL, "Directories", self.headerFormat)
        tmpWsVar.write(0, self.ITEM_COL, "Items", self.headerFormat)
        tmpWsVar.write(0, self.OUTCOME_COL, fixProcedureObject.columnName, self.headerFormat)

        return self.setFixArg(fixProcedureObject, arg)


    def setFixArg(self, fixProcedureObject, arg) -> bool:
        if not fixProcedureObject.validatorFunction:
            pass
        elif ((arg := fixProcedureObject.validatorFunction(arg, fixProcedureObject.argBoundary)) is None):
            return False

        self.fixProcedureArgs[fixProcedureObject] = arg
        return True
    
    
    def processFile(self, longDirAbsolute, dirAbsolute, fileName, needsFolderWritten, tpeIndex):
        countAsError = False
        
        # Onenote files have a ".one-----" extension. The longest onenote extension is 8 characters long. Ignore them.
        # > Technically, something called "fileName.one.txt" would get ignored, but the likelihood of that existing is very low
        # Hidden files should be ignored. This includes temporary Microsoft files (begins with "~$").
        if fileName.startswith("~$") or ".one" in fileName[-8:]:
            return
    
        longFileAbsolute = longDirAbsolute + "\\" + fileName

        hiddenFileSkipStatus = self.hiddenFileCheck(longFileAbsolute)
        if hiddenFileSkipStatus == 2:
            return

        ## THREADING STUFF STARTS
        futures = {
            self.fileProcedureThreadPoolExecutors[tpeIndex].submit(
                findProcedureObject.mainFunction,
                longFileAbsolute,
                dirAbsolute,
                fileName,
                self.findSheets[findProcedureObject]
            ): self.findSheets[findProcedureObject]
            for findProcedureObject in self.fileFindProcedures
        }

        # If not hidden, append to futures with fixProcedureObjects 
        if hiddenFileSkipStatus == 0:
            for fixProcedureObject in self.fileFixProcedures:
                futures[self.fileProcedureThreadPoolExecutors[tpeIndex].submit(
                    self.fixProcedureFunctions[fixProcedureObject],
                    longFileAbsolute,
                    longDirAbsolute,
                    dirAbsolute,
                    fileName,
                    self.fixSheets[fixProcedureObject],
                    self.fixProcedureArgs[fixProcedureObject]
                )] = self.fixSheets[fixProcedureObject]

        for fut in as_completed(futures):
            result = fut.result()
            status = result[0]
            fileSheet = futures[fut]

            if (status == True):
                with self.lockFileNeedsFolderWritten:
                    needsFolderWritten.add(fileSheet)
                countAsError = True
            elif (not status):  # returning None or False
                continue  # It's not super necessary to continue here, but might as well
            elif (status == 2):  # Special case (ex: Used by List All Files)
                with self.lockFileNeedsFolderWritten:
                    needsFolderWritten.add(fileSheet)
            elif (status == 3):  # Special case (ex: used by Identical File)
                countAsError = True

            with self.workbookLock:
                for ewp in result[1:]:
                    ewp.executeWrite()

        # So two files don't finish and try to increment these counters simultaneously
        with self.lockFileScan:
            filesScannedSharedVar.FILES_SCANNED += 1
            self.filesScannedCount += 1
            if countAsError:
                self.fileErrorCount += 1


    def fileCrawl(self, longDirAbsolute, dirAbsolute, dirFiles: List[str]):
        needsFolderWritten = set()

        futures = []
        for fileName in dirFiles:
            with self.fileThreadsCounterLock:
                nextFileThreadIndex = self.fileThreadsCounter % self.numFileThreads
                self.fileThreadsCounter += 1
                
            futures.append(self.fileThreadPoolExecutor.submit(
                self.processFile,
                longDirAbsolute,
                dirAbsolute,
                fileName,
                needsFolderWritten,
                nextFileThreadIndex
            ))

        # If any thread raises an exception, this will ensure they are raised in this "main" WorkbookManager thread
        # , as opposed to just using concurrent.futures(wait)
        for fut in as_completed(futures):
            fut.result()

        return needsFolderWritten            
    
    #def processFolder(self, ...):
    #    pass

    def folderCrawl(self, dirAbsolute, dirFolders, dirFiles):
        needsFolderWritten = set()
        countAsError = False

        futures = {}

        for findProcedureObject in self.folderFindProcedures:
            futures[self.folderProcedureThreadPoolExecutor.submit(
                findProcedureObject.mainFunction,
                dirAbsolute,
                dirFolders,
                dirFiles,
                self.findSheets[findProcedureObject]
            )] = self.findSheets[findProcedureObject]

        for fixProcedureObject in self.folderFixProcedures:
            futures[self.folderProcedureThreadPoolExecutor.submit(
                self.fixProcedureFunctions[fixProcedureObject],
                dirAbsolute,
                dirFolders,
                dirFiles,
                self.fixSheets[fixProcedureObject],
                self.fixProcedureArgs[fixProcedureObject]
            )] = self.fixSheets[fixProcedureObject]

        for fut in as_completed(futures):
            result = fut.result()
            status = result[0]
            folderSheet = futures[fut]

            if status == True:
                with self.lockFolderNeedsFolderWritten:
                    needsFolderWritten.add(folderSheet)
                countAsError = True
            elif not status:  # returning None or False
                pass
            elif status == 2:
                with self.lockFolderNeedsFolderWritten:
                    needsFolderWritten.add(self.fixSheets[fixProcedureObject])
            elif status == 3:
                countAsError = True

            with self.workbookLock:
                for ewp in result[1:]:
                    ewp.executeWrite()


        # NOTE: if creating another layer of threading, need some sort of lock here
        self.foldersScannedCount += 1
        if countAsError:
            self.folderErrorCount += 1

        return needsFolderWritten


    def isHidden(self, longItemAbsolute):
        return bool(os.stat(longItemAbsolute).st_file_attributes & stat.FILE_ATTRIBUTE_HIDDEN)

    # When we want to INCLUDE all hidden files (for find procedures only)
    def includeHiddenFilesCheck(self, longFileAbsolute):
        # 0 == not hidden
        # 1 == hidden, but not required to skip for find procedures, just for fix procedures
        if self.isHidden(longFileAbsolute):
            return 1
        else:
            return 0
    
    # When we want to EXCLUDE all hidden files
    def excludeHiddenFilesCheck(self, longFileAbsolute):
        # 0 == not hidden
        # 2 == hidden, and we must skip in totality
        if self.isHidden(longFileAbsolute):
            return 2
        else:
            return 0
        

    def doFileProceduresExist(self):
        return (len(self.fileFindProcedures) + len(self.fileFixProcedures)) != 0

    def doFolderProceduresExist(self):
        return (len(self.folderFindProcedures) + len(self.folderFixProcedures)) != 0
    

    def initFileCrawlOnly(self, longDirAbsolute, dirAbsolute, dirFiles, _):
        return self.fileCrawl(longDirAbsolute, dirAbsolute, dirFiles)

    def initFolderCrawlOnly(self, _, dirAbsolute, dirFiles, dirFolders):
        return self.folderCrawl(dirAbsolute, dirFolders, dirFiles)

    def initFileFolderCrawl(self, longDirAbsolute, dirAbsolute, dirFiles, dirFolders):
        fileCrawlFuture = self.crawlThreadPoolExecutor.submit(
            self.fileCrawl,
            longDirAbsolute,
            dirAbsolute,
            dirFiles
        )

        folderCrawlFuture = self.crawlThreadPoolExecutor.submit(
            self.folderCrawl,
            dirAbsolute,
            dirFolders,
            dirFiles
        )
        
        return folderCrawlFuture.result() | fileCrawlFuture.result()
    

    def createFileThreads(self):
        ## Create thread pool executors and necessary locks
        numFileProcedures = len(self.fileFindProcedures) + len(self.fileFixProcedures)

        self.fileProcedureThreadPoolExecutors = []

        # Dynamically choose the number of file threads based on a hard-coded max total.
        totalThreads = 150
        self.numFileThreads = totalThreads // numFileProcedures

        self.fileThreadsCounter = 0  # used by self.fileCrawl()
        self.fileThreadsCounterLock = Lock()

        for _ in range(self.numFileThreads):
            self.fileProcedureThreadPoolExecutors.append(
                ThreadPoolExecutor(max_workers = numFileProcedures)
            )

        self.fileThreadPoolExecutor = ThreadPoolExecutor(max_workers = self.numFileThreads)

        self.lockFileNeedsFolderWritten = Lock()
        self.lockFileScan = Lock()

    def createFolderThreads(self):
        numFolderProcedures = len(self.folderFindProcedures) + len(self.folderFixProcedures)
        self.folderProcedureThreadPoolExecutor = ThreadPoolExecutor(max_workers = numFolderProcedures)

        self.lockFolderNeedsFolderWritten = Lock()

    def createCrawlThreads(self):
        # One for file crawl and one for folder crawl
        self.crawlThreadPoolExecutor = ThreadPoolExecutor(max_workers = 2)

    def createSheetLocks(self):
        self.sheetLocks = {}
        for ws in self.getAllProcedureSheets():
            self.sheetLocks[ws] = Lock()


    def initiateCrawl(self, baseDirAbsolute, includeSubfolders, allowModify, includeHiddenFiles, addRecommendations, excludedDirs):
        def addLongPathPrefix(dirAbsolute):
            if dirAbsolute.startswith('\\\\'):
                return '\\\\?\\UNC' + dirAbsolute[1:]
            else:
                return '\\\\?\\' + dirAbsolute
        
        start = time()

        if includeHiddenFiles:
            self.hiddenFileCheck = lambda longFileAbsolute: self.includeHiddenFilesCheck(longFileAbsolute)
        else:
            self.hiddenFileCheck = lambda longFileAbsolute: self.excludeHiddenFilesCheck(longFileAbsolute)
        
        self.styleSummarySheet(baseDirAbsolute, includeSubfolders, allowModify, includeHiddenFiles, addRecommendations)
        self.excludedDirs = excludedDirs
        

        if self.doFileProceduresExist():
            self.createFileThreads()

            if self.doFolderProceduresExist():
                self.createFolderThreads()
                self.createCrawlThreads()
                crawlFunction = self.initFileFolderCrawl
            else:
                crawlFunction = self.initFileCrawlOnly

        elif self.doFolderProceduresExist():
            self.createFolderThreads()
            crawlFunction = self.initFolderCrawlOnly
        else:
            raise Exception("No procedures selected.")
        
        self.createSheetLocks()

        # Only one thread can access the workbook at a time, hence a lock
        self.workbookLock = Lock()

        #
        sheetsSansNonConcurrent = []
        for findProcedureObject in list(self.findSheets.keys()):
            if findProcedureObject.isConcurrentOnly:
                sheetsSansNonConcurrent.append(self.findSheets[findProcedureObject])
        sheetsSansNonConcurrent.extend(list(self.fixSheets.values()))  # Adds all fix procedures indiscriminately

        ##
        walkObject = []

        if (includeSubfolders):
            walkObject = os.walk(baseDirAbsolute)
        
        else:
            # mimic os.walk()'s output but only for the current directory
            dirFolders = []
            dirFiles = []
            
            for item in os.listdir(baseDirAbsolute):
                if os.path.isfile(os.path.join(baseDirAbsolute, item)):
                    dirFiles.append(item)
                else:
                    dirFolders.append(item)

            walkObject = [(baseDirAbsolute, dirFolders, dirFiles)]
        ##

        ###
        for findProcedureObject in self.findSheets.keys():
            if findProcedureObject.startFunction:
                findProcedureObject.startFunction(self.findSheets[findProcedureObject])

        for fixProcedureObject in self.fixSheets.keys():
           if fixProcedureObject.startFunction:
               fixProcedureObject.startFunction(self.fixProcedureArgs[fixProcedureObject], self.fixSheets[fixProcedureObject])
        ###
        
        # TODO: Clean this up?
        excludedDirsSet = set(excludedDirs)

        #
        initialRows = {}
        for (dirAbsolute, dirFolders, dirFiles) in walkObject:
            # Ignore specifically OneNote_RecycleBin folders. Assumes these NEVER have subfolders.
            # Ignore excluded directories.
            # Get "long file path". If folder is hidden, ignore it. Anything within a hidden folder is inadvertently ignored.
            if (dirAbsolute.endswith("OneNote_RecycleBin")) or (dirAbsolute in excludedDirsSet) or self.isHidden(longDirAbsolute := addLongPathPrefix(dirAbsolute)):
                dirFolders[:] = []
                continue

            initialRows.clear()
            for ws in sheetsSansNonConcurrent:
                initialRows[ws] = self.sheetRows[ws] + 1  # CHANGE HERE TO GET NEXT AVAILABLE ROW

            # union operator usage, lol
            # needsFolderWritten = self.fileCrawl(longDirAbsolute, dirAbsolute, dirFiles) | self.folderCrawl(dirAbsolute, dirFolders, dirFiles)
            needsFolderWritten = crawlFunction(longDirAbsolute, dirAbsolute, dirFiles, dirFolders)

            for ws in needsFolderWritten:
                ws.write(initialRows[ws], self.DIR_COL, dirAbsolute, self.dirFormat)


        for findProcedureObject in self.findSheets.keys():
            if findProcedureObject.postFunction:
                findProcedureObject.postFunction(self.findSheets[findProcedureObject])

        ###
        # TODO: Change this (and everywhere else where applicable) to just call a procedureObject function that returns (ala, a getter) the proper function so that it is encapsulated
        if addRecommendations:
            for fixProcedureObject in self.fixSheets.keys():
                if fixProcedureObject.recommendPostFunction:
                    fixProcedureObject.recommendPostFunction(self.fixSheets[fixProcedureObject])
                elif fixProcedureObject.postFunction:
                    fixProcedureObject.postFunction(self.fixSheets[fixProcedureObject])
        else:
            for fixProcedureObject in self.fixSheets.keys():
                if fixProcedureObject.postFunction:
                    fixProcedureObject.postFunction(self.fixSheets[fixProcedureObject])
        ###

        # shutdown threads
        if self.doFileProceduresExist():
            self.fileThreadPoolExecutor.shutdown(wait=True)
            for i in range(self.numFileThreads):
                self.fileProcedureThreadPoolExecutors[i].shutdown(wait=True)

            if self.doFolderProceduresExist():
                self.folderProcedureThreadPoolExecutor.shutdown(wait=True)
                self.crawlThreadPoolExecutor.shutdown(wait=True)

        elif self.doFolderProceduresExist():
            self.folderProcedureThreadPoolExecutor.shutdown(wait=True)

        self.executionTime = time() - start


    def incrementRow(self, ws):
        with self.sheetLocks[ws]:
            self.sheetRows[ws] += 1

    def incrementFileCount(self, ws):
        with self.sheetLocks[ws]:
            self.summaryCounts[ws] += 1

    def incrementRowAndFileCount(self, ws):
       with self.sheetLocks[ws]:
            self.sheetRows[ws] += 1
            self.summaryCounts[ws] += 1


    def styleSummarySheet(self, dirAbsolute, includeSubFolders, allowModify, includeHiddenFiles, addRecommendations):
        self.summarySheet.set_column(0, 0, 34)
        self.summarySheet.set_column(1, 1, 15)
        
        self.summarySheet.write(0, 0, "Directory Path", self.headerFormat)
        self.summarySheet.write(1, 0, "Excluded Directories", self.headerFormat)
        self.summarySheet.write(2, 0, "Include Subdirectories", self.headerFormat)
        self.summarySheet.write(3, 0, "Allow Modify", self.headerFormat)
        self.summarySheet.write(4, 0, "Include Hidden Files", self.headerFormat)
        self.summarySheet.write(5, 0, "Add Recommendations", self.headerFormat)
        self.summarySheet.write(6, 0, "Argument(s)", self.headerFormat)
        self.summarySheet.write(8, 0, "Directory count", self.headerFormat)
        self.summarySheet.write(9, 0, "Directory error count / %", self.headerFormat)
        self.summarySheet.write(10, 0, "File count", self.headerFormat)
        self.summarySheet.write(11, 0, "File error count / %", self.headerFormat)
        self.summarySheet.write(12, 0, "Execution time (s)", self.headerFormat)

        self.summarySheet.write(0, 1, dirAbsolute, self.summaryValueFormat)
        self.summarySheet.write(2, 1, str(includeSubFolders), self.summaryValueFormat)
        self.summarySheet.write(3, 1, str(allowModify), self.summaryValueFormat)
        self.summarySheet.write(4, 1, str(includeHiddenFiles), self.summaryValueFormat)
        self.summarySheet.write(5, 1, str(addRecommendations), self.summaryValueFormat)


    def populateSummarySheet(self):
        col = 1
        for fixProcedureObject in self.fixProcedureArgs.keys():
            arg = self.fixProcedureArgs[fixProcedureObject]
            
            if arg == None:
                continue
            
            self.summarySheet.write(6, col, f"{arg[0] if len(arg) <= 1 else arg} : {fixProcedureObject.name}", self.summaryValueFormat)
            col += 1

        if self.filesScannedCount == 0: fileErrorPercentage = 0
        else: fileErrorPercentage = round(self.fileErrorCount / self.filesScannedCount * 100, 2)
        if self.foldersScannedCount == 0: folderErrorPercentage = 0
        else: folderErrorPercentage = round(self.folderErrorCount / self.foldersScannedCount * 100, 2)

        self.summarySheet.write_number(8, 1, self.foldersScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(9, 1, self.folderErrorCount, self.summaryValueFormat)
        self.summarySheet.write(9, 2, "{}%".format(folderErrorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(10, 1, self.filesScannedCount, self.summaryValueFormat)
        self.summarySheet.write_number(11, 1, self.fileErrorCount, self.summaryValueFormat)
        self.summarySheet.write(11, 2, "{}%".format(fileErrorPercentage), self.summaryValueFormat)
        self.summarySheet.write_number(12, 1, round(self.executionTime, 4), self.summaryValueFormat)
        
        i = 1
        for exDir in self.excludedDirs:
            self.summarySheet.write_string(1, i, exDir, self.summaryValueFormat)
            i += 1

        i = 0
        for ws in self.getAllProcedureSheets():
            self.summarySheet.write(self.sheetRows[self.summarySheet] + i, 1, self.summaryCounts[ws], self.summaryValueFormat)
            i += 1



    def autofitSheets(self):
        for findProcedureSheet in self.findSheets.values():
            findProcedureSheet.autofit()
        for fixProcedureSheet in self.fixSheets.values():
            fixProcedureSheet.autofit()


    def close(self):
        self.populateSummarySheet()
        self.autofitSheets()
        self.wb.close()
