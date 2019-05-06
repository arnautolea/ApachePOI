# ApachePOI

During parsing of created XML or JSON write values in xslx file
Each dep -> separate sheet with dep name and id in sheet’s name
Each sheet contains headers (first row with bold and centered text) for next columns: Emp ID, Lastname, Firstname, Birthdate (*optionally - using date formatter after casting to Date object in Java), Manager ID, Skills (multiline cell)
Every employee is placed in specified department in one separate row
