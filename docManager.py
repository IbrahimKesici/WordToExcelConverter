from docx import Document

class DocManager():

    def __init__(self, fileName, headers):
        self.doc = Document(fileName)
        self.paragraphs = [paragraph for paragraph in self.doc.paragraphs if paragraph.text != ""]
        self.fileName = fileName
        self.headers = headers
        

    def getDetails(self):
        countryCode = self.fileName.split("_")[-4]
        year = self.fileName.split("_")[-2]
  
        for i in range(0, len(self.paragraphs)):
            if self.paragraphs[i].text.lower() == "overall evaluation":
                textSplited = self.paragraphs[i+1].text.split(", ")
                cityName = textSplited[0]
                countryName = textSplited[1]
                break

        return cityName, countryName, countryCode, year

    def getDescription(self):
    
        for i in range(0, len(self.paragraphs)):
            if self.paragraphs[i].text.lower() == "factor ratings" or self.paragraphs[i].text.lower() == "\nfactor ratings":
                index = i + 3
                break

        description ={}
        for header in self.headers[5:]:
            description[header] = self.paragraphs[index].text
            index += 1

        return description

    