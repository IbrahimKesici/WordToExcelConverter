from docManager import DocManager as docM
import excelManager as excel
from IOManager import IOM


def main():
    headers = ["Year", "City", "City_code", "Country", "Country Code",
               "Housing", "Climate and Physical Conditions", "Pollution", "Disease and Sanitation", "Medical Facilities",
               "Education Facilities", "Infrastructure", "Physical Remoteness", "Political Violence and Repression", "Political and Social Environment",
               "Crime", "Communications", "Cultural and Recreation Facilities", "Availability of Goods and Services"]

    IO = IOM()
    docFiles = IO.getFiles()
    for docFile in docFiles:
        doc = docM(fileName = docFile, headers = headers)

        cityName, countryName, countryCode, year  = doc.getDetails()
        description = doc.getDescription()

        excel.convertToExcel(cityName,"x", countryName, countryCode, year, headers, description)

        print(f"{cityName} - {countryName} - {countryCode} - {year}")
        for key, value in description.items():
            print(f"{key}: {value}")
            print("")

if __name__ == "__main__":
    main()