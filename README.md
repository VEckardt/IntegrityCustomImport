# IntegrityCustomImport
Integrity Gateway extension to create complete Integrity Documents from Excel files

## Use Cases
- Date migration from other Requirement tools via Excel into Integrity LM

## Possible Import Layout
![CustomImport](doc/ExcelImport.PNG)

## Installation
- must be installed locally, because the Integrity Gateway is a local application

## Hints
- It's possible to set the Document ID, then the loader will update an existing document in Integrity
  But: be very careful, because the tool was not intended to offer this feature (even if it works like that)!
- The tool provides a Java loader for Excel, linked into the Integrity Gateway 

## Important
- Please try it out in a test environment first, NEVER go directly with this into production 
- This is NOT intended for a permanent use of Excel together with Integrity
- Use the Standard Excel Integration instead (Integrity Add On, available from the PTC Software Download Page) 

## Documentation
see doc/Technical_Documentation_Migrate_Documents_with_Excel.docx
