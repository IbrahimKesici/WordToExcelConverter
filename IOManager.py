import os



class IOM():

    def __init__(self):
        self.path = os.getcwd() + "\\"
        self.folders = ("Start","Archived", "Completed", "Failed")

        self.__createFolder()

    def __createFolder(self):
        
        for folder in self.folders:
            folderPath = self.path + folder
            if not os.path.exists(folderPath):
                os.makedirs(folderPath)

    def getFiles(self):
        filesPath = self.path + self.folders[0]
        
        docFiles = [file for file in os.listdir(filesPath) if ".docx" in file and not "~" in file]
    
        return docFiles


    def moveToFolder(self, folder):
        pass