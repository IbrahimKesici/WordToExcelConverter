import pandas as pd
from openpyxl.workbook import Workbook
import logging

def convertToExcel(path, results, headers, year):
    dfGroup = pd.DataFrame(columns= headers)

    #Group DataFrame
    for city in results:
        try:
            dfNew = pd.DataFrame([results[city]["Properties"] + results[city]["Content"]], columns = headers)
            dfGroup = dfGroup.append(dfNew, ignore_index=True)
        except Exception as ex:
            logging.error(f'Group_Data, {city} - {ex.__str__()}')
            continue

    headers.insert(5, "Region")
    headers.append("Total Index")
    headers.append("Hardship Premium")
    dfLER = pd.DataFrame(columns= headers)

    #LER DataFrame
    for city in results:
        try:
            dfNew = pd.DataFrame([results[city]["Properties"] + results[city]["Rating"]], columns = headers)
            dfLER = dfLER.append(dfNew, ignore_index=True)
        except Exception as ex:
            logging.error(f'LER_Data, {city} - {ex.__str__()}')
            continue

    #Write to excel
    workbookName = 'LERdata_' + year + '.xlsx'
    destinationPath = path + '\\Completed\\' + workbookName
    with pd.ExcelWriter(destinationPath) as writer:
        dfLER.to_excel(writer, sheet_name='LER_Data', index= False)  
        dfGroup.to_excel(writer, sheet_name='Group_Data', index= False)



