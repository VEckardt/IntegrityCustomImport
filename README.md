# IntegrityCustomImport
Integrity Gateway extension to create complete Integrity Lifecycle Manager Documents from Excel files
This tool provides a Java loader for Excel, linked into the Integrity Gateway 

## Use Cases
- Date migration from other Requirement tools via Excel into Integrity LM

## Possible Import Layout
![CustomImport](doc/ExcelImport.PNG)

## Hints
- It's possible to set the Document ID, then the loader will update an existing document in Integrity
- But: be very careful, because the tool was not intended to offer this feature (even if it works like that)!
- The column External ID is required. This is a reference column, allowing the loader to connect Excel Data rows with Integrity Data rows. Internally the External ID can be mapped to a different field if needed.

## Important
- Please try it out in a test environment first, NEVER go directly with this into production 
- This is NOT intended for a permanent use of Excel together with Integrity
- Use the Standard Excel Integration instead (Integrity Add On, available from the PTC Software Download Page) 

## Documentation & Installation
- must be installed locally, because the Integrity Gateway is a local application
- for detailed instructions, please see doc/Technical_Documentation_Migrate_Documents_with_Excel.docx
