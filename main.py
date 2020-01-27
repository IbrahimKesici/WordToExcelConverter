from docManager import DocManager as docM
import excelManager as excel
from IOManager import IOM
import json 


with open("config\cities.json", mode="r") as f:
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

    IO = IOM()
    docFiles = IO.getFiles()
    resultDict = {}
    for docFile in docFiles:
        doc = docM(fileName = docFile, headers = headers)

        cityName, countryName, countryCode, year  = doc.getDetails()
        cityCode = getMapping(cityName, "Code")
        region = getMapping(cityName, "Region")
        

        ratings = doc.getRating()

        description = doc.getDescription(startIndex = 3)

        resultDict[cityName] = {}
        resultDict[cityName]["Properties"] = [year, cityName, cityCode , countryName, countryCode]
        resultDict[cityName]["Content"] = description
        resultDict[cityName]["Rating"] = [region] + ratings
    
   

    excel.convertToExcel(resultDict, headers)
    

if __name__ == "__main__":
    main()
