/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ptc.services.gateway.importer.excel;

import com.mks.api.response.APIException;
import com.mks.gateway.GatewayLogger;
import com.mks.gateway.data.ExternalItem;
import com.mks.gateway.mapper.ItemMapperException;
import com.mks.gateway.mapper.UnsupportedPrototypeException;
import com.mks.gateway.mapper.config.ItemMapperConfig;
import com.mks.gateway.tool.exception.GatewayException;
import com.ptc.services.gateway.importer.LogAndDebug;
import static com.ptc.services.gateway.importer.LogAndDebug.log;
import com.ptc.services.gateway.importer.MappingConfig;
import static com.ptc.services.gateway.importer.MappingConfig.listMappingConfig;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import static java.lang.System.exit;
import static java.lang.System.getProperty;
import static java.lang.System.out;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author veckardt
 */
public class ExcelParser extends ExcelParserGlobals {

    static ItemMapperConfig mappingConfig;
    static IGatewayDriver iGatewayDriver;
    static String buildRelationship = "";
    static String headerRowNum = "";

    @Override
    public List transformToItems(File sourceFile, Map<String, String> defaultItemFields) throws GatewayException {

        log("* ********************************************************************* *", 1);
        log("* Enhanced Excel Import via Gateway, (c) 2017 PTC Inc.", 1);
        log("* ********************************************************************* *", 1);
        log("* Class Name:              " + this.getClass().getName(), 1);
        log("* Default Locale:          " + Locale.getDefault(), 1);
        log("* Temp Directory:          " + getTempDirectory().getAbsolutePath(), 1);
        log("* mks.gateway.configdir:   " + getProperty("mks.gateway.configdir"), 1);
        log("* sourceFile:              " + sourceFile.getAbsolutePath(), 1);
        log("* configProperty:          " + getConfigParam("config"), 1);
        log("* ********************************************************************* *", 1);

        initParams();
        imSession = getSession();
        iGatewayDriver = getDriver();
        mappingConfig = getDriver().getMappingConfiguration();

        listMappingConfig(mappingConfig, "default");
        Properties configProps = getConfigProperties();

        log("Listing local configuration ...", 1);
        log("ConfigProperties Count: " + configProps.size(), 2);

        // List all Gateway Config Properties
        for (Object key : configProps.entrySet()) {
            log("ConfigProperty: " + key.toString(), 2);
            // or explicitely: log("ConfigProperty: " + super.getConfigParam("config"), 5);
        }
        log("End of Listing local configuration.", 1);

        buildRelationship = getConfigParameter(configProps, "buildRelationship");
        headerRowNum = getConfigParameter(configProps, "parser.headerRow");

        // to be able to upload attachments
        setCurrentDirectory(sourceFile.getParent());

        List<ExternalItem> gatewayItems = new ArrayList<>(1);
        try {
            debug("DefaultItemFields count: " + defaultItemFields.size());
            for (String key : defaultItemFields.keySet()) {
                debug("  defaultItemField: " + key + " = " + defaultItemFields.get(key));
            }

            // This reads in the Workbooks specified in arg[0]
            Workbook sourceWb = null;
            InputStream inputStream = null;
            try {
                inputStream = new FileInputStream(sourceFile);
                sourceWb = WorkbookFactory.create(inputStream);
            } catch (IOException | InvalidFormatException e) {
                throw new GatewayException(e);
            }

            int sheetNumber = Integer.parseInt(pSheetNumber) - 1;
            if (sheetNumber >= sourceWb.getNumberOfSheets() || sheetNumber < 0) {
                result.add("ERROR: The Sheet Number specified: " + (sheetNumber + 1) + " is invalid. The SpreadSheet contains " + sourceWb.getNumberOfSheets()
                        + " sheets. Please change the " + SHEET_NUMBER_PROP + " to be the correct Sheet Number row.");
                displayFinalResult(result);
                exit(-6);
            }
            Sheet sheet = sourceWb.getSheetAt(sheetNumber);
            List<Row> allRows = getAllRows(sheet);

            int headerRow = getHeaderRow(allRows, sheet.getLastRowNum(), sheet.getSheetName(), headerRowNum); // What row contains the header information
            HashMap<Integer, Field> headingsMap = getHeadingsMap(allRows.get(headerRow), pSectionColumnName);

            List<Content> contents = getContents(mappingConfig, sourceFile, allRows, headerRow, headingsMap, pSectionColumnName);
            try {
                inputStream.close();
            } catch (IOException e1) {
                throw new GatewayException(e1);
            } // No longer need the file handle on the Excel Spreadsheet

            checkContentExternalID(mappingConfig, contents, pContentLinkField, pSectionColumnName);

            // VE: List<Content> finalContent = getFinalHierarchy(contents, positionColumnName, p);
            if (DEBUG) {
                print(contents);
            }
            gatewayItems = createIIFList(mappingConfig, contents, sourceFile.getName(), defaultItemFields);
            LogAndDebug wed = new LogAndDebug(true);
            wed.writeIIFtoDisk(gatewayItems, "ExcelParser", "final");

        } catch (UnsupportedPrototypeException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, ex.getMessage(), ex);
            throw new GatewayException(ex);
        } catch (ItemMapperException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, ex.getMessage(), ex);
            throw new GatewayException(ex);
        } catch (APIException ex) {
            Logger.getLogger(ExcelParser.class.getName()).log(Level.SEVERE, ex.getMessage(), ex);
            throw new GatewayException(ex);
        }
        log("* ********************************************************************* *", 1);
        log("* END OF: Enhanced Excel Import via Gateway", 1);
        log("* ********************************************************************* *", 1);
        return gatewayItems;
    }

    /**
     * Returns all Rows from the Excel Sheet
     *
     * @param sheet
     * @return
     */
    private static List<Row> getAllRows(Sheet sheet) {
        List<Row> allRows = new ArrayList<>();
        for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
            allRows.add(sheet.getRow(i));
        }
        return allRows;
    }

    /**
     * Checks the Content's ExternalID and Section Level
     *
     * @param mappingConfig
     * @param finalContent the content to check
     * @param pContentLinkField the link field, usually External ID
     * @param pPositionColumnName the position field, usually the Section
     * @throws com.mks.gateway.mapper.ItemMapperException
     */
    public void checkContentExternalID(ItemMapperConfig mappingConfig, List<Content> finalContent, String pContentLinkField, String pPositionColumnName) throws ItemMapperException {

        mappingConfig = MappingConfig.getContentMappingConfig(mappingConfig);
        Map<String, Integer> ExternalIDMap = new TreeMap<>();
        Content oldContent = null;

        log("Checking content ...", 1);

        for (Content content : finalContent) {

            if (content != null && content.getFieldValuesMap().size() > 0) {
                log("Working in content from row " + content.getRowNum() + " ...", 2);
                if (oldContent != null) {
                    // log("content.getSectionLevel():" + content.getSectionLevel(), 1);
                    // log("oldContent.getSectionLevel():" + oldContent.getSectionLevel(), 2);
                    if (content.getSectionLevel() > oldContent.getSectionLevel() + 1) {
                        result.add("ERROR: The " + pPositionColumnName + " level in row " + content.getRowNum() + " should not be higher than " + (oldContent.getSectionLevel() + 1) + " levels !");
                        displayFinalResult(result);
                        exit(-2);
                    }

                }
                oldContent = content;
                String externalID = content.getExternalID(mappingConfig);

                if (externalID.isEmpty()) {

                    result.add("ERROR: Excel row " + content.getRowNum() + " should not have an empty " + pContentLinkField + "!");
                    displayFinalResult(result);
                    exit(-2);
                } else if (content.getPosition() == null || content.getPosition().trim().isEmpty()) {
                    result.add("ERROR: The field '" + pPositionColumnName + "' in Excel row " + content.getRowNum() + " should not be empty!");
                    displayFinalResult(result);
                    exit(-2);
                } else {
                    if (ExternalIDMap.containsKey(externalID)) {
                        result.add("ERROR: Excel row " + content.getRowNum() + " should not have the same " + pContentLinkField + " (" + externalID + ") like in row " + ExternalIDMap.get(externalID) + "!");
                        displayFinalResult(result);
                        exit(-2);
                    }
                    ExternalIDMap.put(externalID, content.getRowNum());
                }
            }
        }
        log("Checking content done.", 1);
    }

    /**
     * This method parses the Excel file for the contents.
     * <br><br>This method iterates over all the data in an Excel sheet,
     * skipping over the row designated as the header row. It then iterates over
     * the columns by finding only the columns found by
     * {@link #getHeadingsMap(Cell[], String, Properties)} and creates a List of
     * the Content inside the Excel Sheet.
     *
     * @param sheet the sheet to parse
     * @param headerRow the row containing the header information
     * @param headingsMap the map of the headings index to their names
     * @param positionColumnName the column name of the position column
     * @param properties the properties for the Parser
     * @return a list of all the content
     */
    private List<Content> getContents(ItemMapperConfig mappingConfig, File sourceFile, List<Row> allRows, int headerRow, HashMap<Integer, Field> headingsMap,
            String positionColumnName) throws ItemMapperException {
        boolean mapPosition = Boolean.parseBoolean(pMapPosition);
        boolean ignoreBlanks = Boolean.parseBoolean(pIgnoreBlanks);
        int numRows = allRows.size();
        // Need to find the position column for use later on.
        int positionColumnNum = -1;
        Set<Integer> validColumnsIdxs = headingsMap.keySet();
        for (Integer column : validColumnsIdxs) {
            if (headingsMap.get(column).getFieldName().equals(positionColumnName)) {
                positionColumnNum = column;
            }
        }

        // positionColumnNum = 4;
        debug("Position Column Name: " + positionColumnName + ", with positionColumnNum: " + positionColumnNum);

        int startRow = Integer.parseInt(pStartRow) - 1;
        int endRow = Integer.parseInt(pEndRow) - 1;
        // VE: int docIDRow = Integer.parseInt(properties.getProperty(DOCUMENT_ID_ROW, "-1")) - 1;
        endRow = (endRow == -1 ? numRows : Math.min(endRow, numRows));

        debug("Header Row: " + headerRow);
        debug("Starting at: " + startRow + ", Ending at: " + endRow);

        List<Content> contents = new ArrayList<>(numRows);
        for (int i = startRow; i < endRow; i++) {
            debug("Looking at Row " + i); //  + ", docIdRow=" + docIDRow);
            //
            // Part 1: Handle the Values
            //
            if (i > headerRow) {
                HashMap<String, String> fieldValuesMap = new HashMap<>(validColumnsIdxs.size());
                //List<Cell> headerCells = getAllCells(allRows.get(i));
                Row row = allRows.get(i);

                String position = null;
                for (Integer columnId : validColumnsIdxs) {
                    String fieldName = headingsMap.get(columnId).getFieldName();
                    if (row != null && columnId < row.getLastCellNum()/*headerCells.size()*/) { // This prevents the problem where a row isn't as long as a column we think we need.
                        Cell cell = row.getCell(columnId);

                        if (cell != null) {
                            String text = "";

                            if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                                // GatewayLogger.log("In Cell.CELL_TYPE_FORMULA ..");
                                // out.println("Formula is " + cell.getCellFormula());
                                switch (cell.getCachedFormulaResultType()) {
                                    case Cell.CELL_TYPE_NUMERIC:
                                        GatewayLogger.log("Last evaluated as: " + cell.getNumericCellValue());
                                        break;
                                    case Cell.CELL_TYPE_STRING:
                                        // out.println("Last evaluated as \"" + cell.getRichStringCellValue() + "\"");
                                        text = cell.getRichStringCellValue().toString();
                                        // text = "<!-- MKS HTML -->" + text;
                                        break;
                                }
                            } else {

                                // log("IN reading cell ....", 1);
                                DataFormatter dataFormatter = new DataFormatter();
                                if (headingsMap.get(columnId).isRelationship()) {
                                    // cell.setCellType(Cell.CELL_TYPE_STRING);
                                    text = dataFormatter.formatCellValue(cell).replace(".0", "").replace(".", ",");
//
////                                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
////                                    text = Double.toString(cell.getNumericCellValue());
                                } else {
//                                    cell.setCellType(Cell.CELL_TYPE_STRING);
//                                    text = cell.getStringCellValue();
                                    text = dataFormatter.formatCellValue(cell);

                                    // GatewayLogger.log(">>1 " + text);
//                                    if (text.startsWith("<div>")) {
//                                        text = "<!-- MKS HTML -->" + text;
//                                        GatewayLogger.log(">>2");
//                                        GatewayLogger.log(text);
//                                    }
                                }
                            }
                            // Put the Text into the Map if some text is there
                            if (text.trim().length() > 0 || !ignoreBlanks) { // If the text isn't blank OR if blanks don't matter 

                                // text = text.replace("<", "&lt;").replace(">", "&gt;");
                                // replace the charage returns
                                if (text.contains("" + (char) (10))) {
                                    // log("Char 10 found, len=" + text.length(), 2);
                                    text = text.replace("" + (char) (10), "<br>");
                                }

                                // put into the map
                                fieldValuesMap.put(headingsMap.get(columnId).getFieldName(), text);

                                if (positionColumnNum == columnId) { //If this is the position column, update the position
                                    position = text;
                                    if (!mapPosition) // And if we are not mapping the position field, remove it. 
                                    {
                                        fieldValuesMap.remove(headingsMap.get(columnId).getFieldName());
                                    }
                                }
                            }

                            // log("! " + headingsMap.get(columnId).getFieldName() + ", " + fieldValuesMap.get("Category"), 1);
                            // com.mks.gateway.App.main(new String[]{"import"});
                            // if the field name is Text and Category is Heading, then set the Heading format
                            // VE: Deactivated: 2016/11/26                            
                            if (!((getConfigParam("setHeadingFormat") + "").contentEquals("false")) && headingsMap.get(columnId).getFieldName().contentEquals("Text")) {
                                String category = fieldValuesMap.get("Category");
                                // log("category: " + category, 2);
                                // String section = fieldValuesMap.get("Section");
                                // log("position: " + position, 2);
                                if (category != null && category.contentEquals("Heading")) {
                                    int level = position.replaceAll("[^.]", "").length() + 1;
                                    fieldValuesMap.put("Text", "<h" + level + ">" + fieldValuesMap.get("Text") + "</h" + level + ">");
                                }
                            }
                            // load the images into Integrity automatically
                            if (fieldName.contentEquals("Images")) {

                                // image replacement requested
                                //
                                DataFormatter dataFormatter = new DataFormatter();
                                String imageTexts = dataFormatter.formatCellValue(cell);
                                log("Image found: " + dataFormatter.formatCellValue(cell), 2);
                                //

                                if (!imageTexts.trim().isEmpty()) {
                                    for (String imageText : imageTexts.split("\n")) {
                                        // {Ext_ID_5:$Bild1$=beispiel.png, Ext_ID_5:$Bild2$=beispiel2.png}
                                        log("Image Text: " + imageText, 2);
                                        // = $Bild2$:beispiel2.png
                                        String imageTag = imageText.split(":")[0];
                                        log("Image Tag: " + imageTag, 2);
                                        if (fieldValuesMap.get("Text").contains(imageTag)) {
                                            String fileName = imageText.split(":")[1];
                                            log("File Name: " + fileName, 2);
                                            // attachmentList.put(fieldValuesMap.get("External ID") + ":" + imageTag, fileName);

                                            String imagePathAndFile = sourceFile.getParent() + "\\Images\\" + fileName;
                                            File file = new File(imagePathAndFile);
                                            if (!file.exists()) {
                                                String errorMessage = "The image file '" + imagePathAndFile + "' can not be found.\nPlease validate the name and check that the file exists.";
                                                result.add("ERROR: " + errorMessage);
                                                displayFinalResult(result);

                                                exit(-8);
                                            }

                                            // <img src="190/Text%20Attachments/eclipse.jpg" height="120" width="120" />
                                            log("Updating tag '" + imageTag + "' with " + "<img src=\"Images\\" + fileName + "\" />", 2);
                                            fieldValuesMap.put("Text", fieldValuesMap.get("Text").replace(imageTag, "<img src=\"Images\\" + fileName + "\" height=\"260\" width=\"319\"/ />"));
                                        }

                                        // <img height="120" src="mks:///item/field?fieldid=Text Attachments&attachmentname=eclipse.jpg" width="120">
                                    }
                                }
                            }
                        }
                    }
                }

                // Transform code
//                boolean continueTransform = true;
//                int transformIndex = 1;
//                do {
//                    String newName = properties.getProperty(TRANSFORM_FIELDS_PREFIX_PROP + "." + transformIndex + "." + TRANSFORM_FIELDS_NAME_SUFFIX);
//                    if (newName == null) {
//                        continueTransform = false;
//                    } else {
//                        String transformPattern = properties.getProperty(TRANSFORM_FIELDS_PREFIX_PROP + "." + transformIndex + "." + TRANSFORM_FIELDS_PATTERN_SUFFIX);
//                        if (transformPattern == null) {
//                            debug("Transformation for Transform Index " + transformIndex + " missing.");
//                            transformIndex++;
//                            continue;
//                        }
//                        StringBuilder sb = new StringBuilder();
//                        int curIdx = -1;
//                        int startIdx = 0;
//                        while ((curIdx = transformPattern.indexOf("{", startIdx)) >= 0) { //Searches for a pattern
//                            if (curIdx > 0) {
//                                sb.append(transformPattern.substring(startIdx, curIdx));
//                            }
//                            if (curIdx == (transformPattern.length() - 1)) {
//                                sb.append("{");
//                                startIdx = curIdx + 1;
//                            } else {
//                                int endIndex = transformPattern.indexOf("}", curIdx);
//                                if (endIndex < 0) {
//                                    // no matching closing token, don't expand
//                                    break;
//                                }
//                                String fieldName = transformPattern.substring(curIdx + 1, endIndex);
//                                // Fills in the the property with either the value or blank.
//                                sb.append(fieldValuesMap.get(fieldName) != null ? fieldValuesMap.get(fieldName) : "");
//                                fieldValuesMap.remove(fieldName);
//                                startIdx = endIndex + 1;
//                            }
//                            if (sb.toString().trim().length() > 0 || !ignoreBlanks) {
//                                fieldValuesMap.put(newName, sb.toString());
//                            }
//                        }
//                    }
//                    transformIndex++;
//                } while (continueTransform == true); //iterate over all transform properties				
                //If there are no fields to map and we are ignoring blanks, then don't add the content
                if (!(fieldValuesMap.isEmpty() && ignoreBlanks)) {

                    // if (Boolean.parseBoolean(properties.getProperty(IS_NEW_DOC_PROP, "true"))) // If this is a new document, explicitly set the ID to blank
                    // {
                    //     fieldValuesMap.put(properties.getProperty(LINK_FIELD_PROP, "DOCUMENT_ID"), "");
                    // }
                    // if (position == null) {
                    //     contents.add(new Content(fieldValuesMap, i));
                    // } else {
                    contents.add(new Content(position, fieldValuesMap, i));
                    // }
                }

                //
                // Part 2: Handle the Header
                //
            } else if (i < headerRow) {

                Row row = allRows.get(i);
                if (row != null) {
                    String fieldName = getCellValue(row, 0).replace(":", "");
                    String fieldValue = getCellValue(row, 1);
                    // ItemMapperConfig docMappingConfig = mappingConfig.getSubConfig("DOCUMENT");
                    List<ItemMapperConfig.Field> fieldList = mappingConfig.getExternalOutgoingFields("DOCUMENT");
                    for (ItemMapperConfig.Field field : fieldList) {
                        if (field.externalField.equals(fieldName)) {
                            if (field.fieldType.toLowerCase().equals("id")) {
                                fieldValue = fieldValue.replaceAll("\\.0*$", "");
                            }
                        }
                    }
                    docFields.add(new Field(fieldName, fieldValue));
                }
            }
        }
        return contents;
    }

    private static String getCellValue(Row row, int cellId) {
        if (row != null) {
            Cell cell = row.getCell(cellId);
            if (cell != null) {

//                String rsdata = "";
//                try {
//                    rsdata = cell.getStringCellValue();
//                } catch (NumberFormatException ex) {
//                    rsdata = cell.getNumericCellValue() + "";
//                }
                cell.setCellType(Cell.CELL_TYPE_STRING);
                return cell.toString();
            } else {
                return "";
            }
        } else {
            return "";
        }
    }

    /**
     * This parses out the heading names from the headings row. The headings row
     * is determined by {@value #HEADER_ROW_PROP} or is always the top row if
     * that property is unspecified. The Headers become the names of the
     * "external" field names to be used as part of the Gateway mapper XML file.
     * If a Cell's contents in this row is empty it is called "EMPTY_HEADER_X"
     * where X is the column number. <br><br>
     * If the values of {@value #COLUMN_NAMES_PROP} can be found then they are
     * included, if they cannot no error is thrown. If no values are specified
     * then all columns will be used. In all cases the Position column is
     * included in this list, but only included in the exported IIF if
     * {@value #MAP_POSITION_PROP} is <tt>true</tt>.
     *
     * @param headerCells the Array of Cells to search.
     * @param positionColumnName the position column name (if one exists)
     * @param properties the properties of this Parser
     * @return A map that contains the column index mapped to the column name to
     * be used when retrieving the values from the Excel sheet.
     */
    private static HashMap<Integer, Field> getHeadingsMap(Row headerRow, String positionColumnName) throws ItemMapperException {
        String columns = pColumnNames;
        List<Cell> headerCells = getAllCells(headerRow);

        ItemMapperConfig docMappingConfig = MappingConfig.getDocumentMappingConfig(mappingConfig);

        @SuppressWarnings("unchecked")
        List<String> validColumns = new ArrayList<>((columns != null && !columns.equals(""))
                ? (Arrays.asList(columns.split(pColumnDelim)))
                : Collections.EMPTY_LIST);
        if (validColumns.contains(positionColumnName) && positionColumnName != null) {
            validColumns.remove(positionColumnName);
        }
        HashMap<Integer, Field> headingsMap = new HashMap<>();
        debug("Columns specified in properties, not including position column: " + validColumns);
        for (int i = 0; i < headerCells.size(); i++) {
            Cell headingCell = headerCells.get(i);
            String cellContents = headingCell.getStringCellValue();
            if (validColumns.contains(cellContents)
                    || validColumns.isEmpty()
                    || cellContents.equals(positionColumnName)) {
                // This accommodates if the Header column has an empty value
                cellContents = cellContents.isEmpty() ? "EMPTY_HEADER_" + (i + 1) : cellContents;
                debug("Mapping Header " + i + " to " + cellContents);
                headingsMap.put(i, new Field(cellContents, "", MappingConfig.getExternalFieldType(docMappingConfig, cellContents)));
            }
        }
        if (headingsMap.isEmpty()) {
            result.add("No valid columns found. Property: " + COLUMN_NAMES_PROP + " has contents: " + validColumns);
            displayFinalResult(result);
            exit(-6);
        }
        return headingsMap;
    }

    private static List<Cell> getAllCells(Row headerRow) {
        List<Cell> allCells = new ArrayList<>();
        Iterator<Cell> cellIter = headerRow.cellIterator();
        while (cellIter.hasNext()) {
            allCells.add(cellIter.next());
        }
        return allCells;
    }

    /**
     * Returns the value of the Header row after checking to make sure it is an
     * acceptable value.
     * <br><br>The value must be greater than 0 and less than the number of rows
     * in the Excel Sheet. The rows are counted starting from 1 as that is how
     * Excel refers to them. If {@value #HEADER_ROW_PROP} is not specified, then
     * row 1 will be used.
     *
     * @param numRows The total number of rows in the sheet
     * @param sheetName the sheet name to look at
     * @param properties the properties of this Parser
     * @return the integer of the header row.
     *
     */
    private static int getHeaderRow(List<Row> allRows, int numRows, String sheetName, String headerRowNum) {

        // VE: changed to get this value automatically
        int headerRow = -1;

        out.println("headerRowNum: " + headerRowNum);

        if (!headerRowNum.isEmpty()) {
            headerRow = Integer.parseInt(headerRowNum);
            if (headerRow >= numRows) {
                out.println("The header row specified: " + headerRow + " is greater than the number of rows " + numRows
                        + " in sheet '" + sheetName + "'. Please change the " + HEADER_ROW_PROP + " to be the correct header row.");
                exit(-5);
            } else if (headerRow < 1) {
                out.println("The header row specified: " + headerRow + " is invalid"
                        + ". Please change the " + HEADER_ROW_PROP + " to be the correct header row.");
                exit(-5);
            }
            return headerRow - 1;
        } else {

            for (Row row : allRows) {

                headerRow++;
                String cellValueFirstCell = getCellValue(row, 0);
                if (!cellValueFirstCell.isEmpty() && cellValueFirstCell.contentEquals(pSectionFieldName)) {
                    return headerRow;
                }
            }
        }
        String errorMessage = "The header row could not be determined from sheet '" + sheetName + "'.\nPlease make sure that the sheet contains the field '" + pSectionFieldName + "' in your data header row.";
        result.add("ERROR: " + errorMessage);
        displayFinalResult(result);

        exit(-5);
        return headerRow;
    }

    public static void debug(String text) {
        // if (DEBUG) {
        // out.println(text);
        // LOGGER.log(Level.INFO, text);
        GatewayLogger.log(text);
        // }
    }

    public static void displayFinalResult(List<String> result) {
        if (result != null && !result.isEmpty()) {
            StringBuffer message = new StringBuffer();
            message.append("A problem occurred. Please check the input document!\n\n");
            for (String string : result) {
                message.append("").append(string).append("\n");
            }
            GatewayLogger.logMessage(message.toString(), 10);
            //log(message.toString());
            JOptionPane jop = new JOptionPane();
            jop.setToolTipText("Log the output of the export");
            JOptionPane.showMessageDialog(null, message, "Problem with creating an Integrity document", JOptionPane.WARNING_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null, "Integrity Document has been sucessfully created!");
        }
    }

//    Function fnConvert2HTML(myCell As Range) As String
//    Dim bldTagOn, itlTagOn, ulnTagOn, colTagOn As Boolean
//    Dim i, chrCount As Integer
//    Dim chrCol, chrLastCol, htmlTxt As String
//    
//    bldTagOn = False
//    itlTagOn = False
//    ulnTagOn = False
//    colTagOn = False
//    chrCol = "NONE"
//    htmlTxt = ""
//    chrCount = myCell.Characters.Count
//    
//    For i = 1 To chrCount
//        With myCell.Characters(i, 1)
//            If (.Font.Color) Then
//                chrCol = fnGetCol(.Font.Color)
//                If Not colTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    colTagOn = True
//                Else
//                    If chrCol <> chrLastCol Then htmlTxt = htmlTxt & ""
//                End If
//            Else
//                chrCol = "NONE"
//                If colTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    colTagOn = False
//                End If
//            End If
//            chrLastCol = chrCol
//            
//            If .Font.Bold = True Then
//                If Not bldTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    bldTagOn = True
//                End If
//            Else
//                If bldTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    bldTagOn = False
//                End If
//            End If
//    
//            If .Font.Italic = True Then
//                If Not itlTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    itlTagOn = True
//                End If
//            Else
//                If itlTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    itlTagOn = False
//                End If
//            End If
//    
//            If .Font.Underline > 0 Then
//                If Not ulnTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    ulnTagOn = True
//                End If
//            Else
//                If ulnTagOn Then
//                    htmlTxt = htmlTxt & ""
//                    ulnTagOn = False
//                End If
//            End If
//            
//            If (Asc(.Text) = 10) Then
//                htmlTxt = htmlTxt & "
//"
//            Else
//                htmlTxt = htmlTxt & .Text
//            End If
//        End With
//    Next
//    
//    If colTagOn Then
//        htmlTxt = htmlTxt & ""
//        colTagOn = False
//    End If
//    If bldTagOn Then
//        htmlTxt = htmlTxt & ""
//        bldTagOn = False
//    End If
//    If itlTagOn Then
//        htmlTxt = htmlTxt & ""
//        itlTagOn = False
//    End If
//    If ulnTagOn Then
//        htmlTxt = htmlTxt & ""
//        ulnTagOn = False
//    End If
//    htmlTxt = htmlTxt & ""
//    // fnConvert2HTML = htmlTxt
//End Function
//Function fnGetCol(strCol As String) As String   
//    Dim rVal, gVal, bVal As String
//    strCol = Right("000000" & Hex(strCol), 6)
//    bVal = Left(strCol, 2)
//    gVal = Mid(strCol, 3, 2)
//    rVal = Right(strCol, 2)
//    fnGetCol = rVal & gVal & bVal
//End Function
// Read more at http://weijie.info/InMyHead/automatically-convert-formatted-excel-text-into-html/#6RZyHmbeA1V9DdBA.99
    private String getConfigParameter(Properties configProps, String param) {
        String configParam = configProps.getProperty(param);
        return (configParam == null ? "" : configParam);
    }
}
