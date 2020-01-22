from docx import Document


def readDoc(fileName):
    doc = Document(fileName)

    return doc

def getDocDetails(doc):
   
    return cityName, countryName, date, countryCode

def searchHeader(header, doc):
    for table in doc.tables:
        for cell in table.rows[0].cells:
            for paragraph in cell.paragraphs:
                if "title" in paragraph.text:
                    print(paragraph.text)
                    return paragraph.text
        
    

    