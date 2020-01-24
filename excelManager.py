import pandas as pd
from openpyxl.workbook import Workbook

def convertToExcel(cityName, cityCode, countryName, countryCode, year, headers, descriptions):
    
    df = pd.DataFrame(columns= headers)
    descriptionList = []
    
    descriptionList.append(year)
    descriptionList.append(cityName)
    descriptionList.append(cityCode)
    descriptionList.append(countryName)
    descriptionList.append(countryCode)

    for key, value in descriptions.items():
        descriptionList.append(value)

    df2 = pd.DataFrame(descriptionList, columns = headers)
    df.append(df2)
    #pd.concat([pd.DataFrame(descriptionList, columns = headers)], ignore_index=True)

    df.to_excel("output.xlsx", sheet_name= "Group_Data", index= False) 





