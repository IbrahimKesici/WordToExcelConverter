# WordToExcel

Python script to automate the process of reading .docx content: category ratings and descriptions, and generate .xlsx file

    1 - Get docx files from Start Folder( Manually convert doc to docx if you have doc files)
    2 - Get cityName and countryName info from .docx content, countryCode and year info from .docx file name, cityCode and region from cities.json file
    3 - Read ratings from "Overall Evaluation - Factor" table and description from particular paragraphs
    4 - Write final results to .xlsx file as two sheets: LER_Data(ratings) and Group_Data(descriptions)