import docManager as docM
import excelManager
from IOManager import IOM


def main():
    headers = ["", ""]

    IO = IOM()
    docFiles = IO.getFiles()

    for docFile in docFiles:
        docContent = docM.readDoc(docFile)
        cityName, countryName, countryCode, date  = docM.getDocDetails(docContent)

        description = {}
        for header in headers:
            description[header] = docM.searchHeader(header, docContent)
            
if __name__ == "__main__":
    main()