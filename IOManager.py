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

    def convertToDocx(self):
        #TODO: Convert doc files to docx 
        pass

    def getFiles(self):
        filesPath = self.path + self.folders[0]
        
        docFiles = [filesPath + "\\" + file for file in os.listdir(filesPath) if ".docx" in file and not "~" in file]
    
        return docFiles


    def moveToFolder(self, folder):
        pass