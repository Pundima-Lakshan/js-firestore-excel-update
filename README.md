# js-firestore-excel-update
 Update firebase firestore database using data in excel files


## How to
Run the index.js in a node environment

Initialize firebase configuration (Ln 12)

Select the excel file (Ln 26)

- const workbook = XLSX.readFile(````${letter}.xlsx````);

Configure the needed fields and data to update (Ln 27 ...)

Reference a collection in firestore (Ln 32)

- const collectionName = ```"demoGuests"```;

## Dependancies

- xlsx
- firebase