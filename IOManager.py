import os, re, shutil
import logging
import win32com.client as win32
from win32com.client import constants

class IOM():

    def __init__(self):
        self.path = os.getcwd() + "\\" #TODO: Change it to actual directory
        self.folders = ("Start", "Archived", "Completed", "Failed")

        self.__createFolder()

    def __createFolder(self):
        
        for folder in self.folders:
            folderPath = self.path + folder
            if not os.path.exists(folderPath):
                os.makedirs(folderPath)

    def convertDocToDocx(self):
        """ Convert .doc files on specified directory to .docx files
        """
        path = self.path + self.folders[0]
        docFiles = self.__readFiles(path = path, fileFormat='.doc')

        word = win32.gencache.EnsureDispatch('Word.Application')
        for docFile in docFiles:
            try:
                doc = word.Documents.Open(docFile)
                doc.Activate ()

                # Rename path with .docx
                newFilePath = os.path.abspath(docFile)
                newFilePath = re.sub(r'\.\w+$', '.docx', newFilePath)

                # Save and Close
                word.ActiveDocument.SaveAs(newFilePath, FileFormat=constants.wdFormatXMLDocument)
                doc.Close(True)
            except Exception as ex:
                logging.error(f'{docFile} is not converted to .docx  - {ex.__str__()}')
                doc.Close(True)
                continue

         

    def getCWPath(self):
        """ Return to current working directory
        """
        return self.path

    def getFiles(self):
        """ Get .docx files from the specified directory
        """
        filesPath = self.path + self.folders[0]

        docxFiles = self.__readFiles(filesPath)
        return docxFiles

    def moveToFolder(self, filePath, folderIndex = 1):
        """Move files to speficied folder
        param filePath: str, path of .docx file with its name
        param folderIndex: str, index of the folder
        """
        sourcePath = filePath
        destinationPath = self.path + self.folders[folderIndex]

        shutil.move(sourcePath, destinationPath)
        

    def __readFiles(self, path, fileFormat = '.docx'):
        """ Get files on the specified path
        param path: str, path of the files
        param fileFormat: str, format of the file
        """
        files = [path + "\\" + fileName for fileName in os.listdir(path) if fileName.endswith(fileFormat) and not "~" in fileName]
        return files