from docManager import DocManager as docM
import excelManager as excel
from IOManager import IOM
import json, logging


with open("config\\cities.json", mode="r") as f:
    citiesJSON = json.load(f)  


def getMapping(cityName, field):
    try:
        result = citiesJSON.get(cityName)[field]
    except:
        return "NA"

    return result


def main():
    headers = ["Year", "City", "City_code", "Country", "Country Code",
               "Housing", "Climate and Physical Conditions", "Pollution", "Disease and Sanitation", "Medical Facilities",
               "Education Facilities", "Infrastructure", "Physical Remoteness", "Political Violence and Repression", "Political and Social Environment",
               "Crime", "Communications", "Cultural and Recreation Facilities", "Availability of Goods and Services"]
    operationResult = {}
    logging.basicConfig(filename= 'log.txt',
                        level = logging.INFO,
                        datefmt= '%d %b %Y - %H:%M:%S',
                        format = '%(asctime)s: %(funcName)s: %(levelname)s: - %(message)s',
                        filemode= 'a')
    
    IO = IOM()
    IO.convertDocToDocx()
    docxFiles = IO.getFiles()

    resultDict = {}
    for docxFile in docxFiles:
        try:
            doc = docM(fileName = docxFile, headers = headers)

            cityName, countryName, countryCode, year  = doc.getDetails()
            cityCode = getMapping(cityName, "Code")
            region = getMapping(cityName, "Region")
            
            ratings = doc.getRating()
            description = doc.getDescription(startIndex = 3)

            resultDict[cityName] = {}
            resultDict[cityName]["Properties"] = [year, cityName, cityCode , countryName, countryCode]
            resultDict[cityName]["Content"] = description
            resultDict[cityName]["Rating"] = [region] + ratings

            operationResult[docxFile] = 'Success'
        except Exception as ex:
            logging.error(f'{docxFile} - {ex.__str__()}')
            operationResult[docxFile] = 'Failure'
            continue

        
    #Write to Excel    
    try:
        excel.convertToExcel(IO.getCWPath(), resultDict, headers, year)
    except Exception as ex:
        logging.critical(f'Writing to Excel - {ex.__str__()}')
    
    #Move to related folders
   
    for docx, status in operationResult.items():
        try:
            if status == 'Success':
                IO.moveToFolder(filePath = docx, folderIndex = 1)
            else:
                IO.moveToFolder(filePath = docx, folderIndex = 3)
        except Exception as ex:
            logging.warning(f'Process successful, but cannot move file - {ex.__str__()}')
            continue

    
if __name__ == "__main__":
    main()
