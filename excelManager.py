import pandas as pd
from openpyxl.workbook import Workbook

def convertToExcel(results, headers):
    dfGroup = pd.DataFrame(columns= headers)

    #Group DataFrame
    for city in results:
        dfNew = pd.DataFrame([results[city]["Properties"] + results[city]["Content"]], columns = headers)
        dfGroup = dfGroup.append(dfNew, ignore_index=True)

    headers.insert(5, "Region")
    headers.append("Total Index")
    headers.append("Hardship Premium")
    dfLER = pd.DataFrame(columns= headers)

    #LER DataFrame
    for city in results:
        dfNew = pd.DataFrame([results[city]["Properties"] + results[city]["Rating"]], columns = headers)
        dfLER = dfLER.append(dfNew, ignore_index=True)

    #Write to excel
    with pd.ExcelWriter('output.xlsx') as writer:
        dfLER.to_excel(writer, sheet_name='LER_Data', index= False)  
        dfGroup.to_excel(writer, sheet_name='Group_Data', index= False)



