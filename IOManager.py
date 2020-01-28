import os

class IOM():

    def __init__(self):
        self.path = os.getcwd() + "\\" #TODO: Change it to actual directory
        self.folders = ("Start","Archived", "Completed", "Failed")

        self.__createFolder()

    def __createFolder(self):
        
        for folder in self.folders:
            folderPath = self.path + folder
            if not os.path.exists(folderPath):
                os.makedirs(folderPath)

    def getFiles(self):
        filesPath = self.path + self.folders[0]
        docFiles = [filesPath + "\\" + fileName for fileName in os.listdir(filesPath) if ".docx" in fileName and not "~" in fileName]
    
        return docFiles

    def moveToFolder(self, folder):
        pass