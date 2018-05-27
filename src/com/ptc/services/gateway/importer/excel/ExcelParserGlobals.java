/*
 * Copyright:      Copyright 2018 (c) Parametric Technology GmbH
 * Product:        PTC Integrity Lifecycle Manager
 * Author:         V. Eckardt, Principal Solution Architect, ALM
 * Purpose:        Custom Developed Code
 * **************  File Version Details  **************
 * Revision:       $Revision: 1.16 $
 * Last changed:   $Date: 2016/02/12 23:26:57CET $
 */
package com.ptc.services.gateway.importer.excel;

import com.mks.api.response.APIException;
import com.mks.gateway.GatewayLogger;
import com.mks.gateway.data.ExternalItem;
import com.mks.gateway.data.ItemField;
import com.mks.gateway.driver.GatewayTransformer;
import com.mks.gateway.mapper.ItemMapperException;
import com.mks.gateway.mapper.ItemMapperSession;
import com.mks.gateway.mapper.UnsupportedPrototypeException;
import com.mks.gateway.mapper.config.ItemMapperConfig;
import static com.ptc.services.gateway.importer.LogAndDebug.log;
import com.ptc.services.gateway.importer.MappingConfig;
import java.io.File;
import static java.lang.System.exit;
import static java.lang.System.setProperty;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author veckardt
 */
public class ExcelParserGlobals extends GatewayTransformer {

    @Override
    public String getTransformerKey() {
        return "MSEXCEL";
    }

    public static String pColumnDelim = "";
    public static String pColumnNames = "";
    public static String pContentLinkField = "";
    public static String pDebug = "";
    public static String pDocumentLinkField = "";
    public static String pEndRow = "";
    // public static String pHeaderRow = "";
    public static String pSectionFieldName = "";
    public static String pIgnoreBlanks = "";
    public static String pMapPosition = "";
    public static String pSectionColumnName = "";
    public static String pPropertiesFilePath = "";
    public static String pRelationshipFieldName = "";
    public static String pSheetNumber = "";
    public static String pStartRow = "";
    public static String pTransformFieldsNameSuffix = "";
    public static String pTransformFieldsPatternSuffix = "";
    public static String pTransformFieldsPrefix = "";

    public static final List<String> result = new ArrayList<String>();

    // public static final Logger LOGGER = Logger.getLogger(ExcelParser.class.getName());
    // public static Handler fileHandler;
    public static final ArrayList<Field> docFields = new ArrayList<Field>();

    /**
     * The name of the Properties file.
     */
    // public static final String PROPERTIES_FILE_PATH = "excelParser.properties";
    /**
     * The name of the property which specifies the Sheet Number
     */
    public static final String SHEET_NUMBER_PROP = "parser.sheetNumber";

    // public static final String DOCUMENT_ID_ROW = "gateway.documentIDRow";
    /**
     * The name of the property which specifies the Column Names. Default is a
     * comma (",") delimited list. Otherwise delimited by
     * {@link #COLUMN_DELIM_PROP}.
     */
    public static final String COLUMN_NAMES_PROP = "parser.columns";
    /**
     * The name of the property which specifies the "position column" which is
     * the column that designates the position of the content. Used to build
     * hierarchy.
     */
    public static final String SECTION_COLUMN_NAME_PROP = "parser.section.column.name";
    /**
     * The name of the property which specifies if the position column should
     * also be made available for mapping. In some cases the Integrity "Section"
     * field replaces this, so it isn't needed to be mapped. Default false.
     */
    public static final String MAP_POSITION_PROP = "parser.position.isMapped";
    /**
     * The name of the property which specifies if debug logging should be on.
     * Default false.
     */
    public static final String DEBUG_PROP = "parser.debug";
    /**
     * The name of the property which specifies which row contains the header
     * information. Default 1.
     */
    public static final String HEADER_ROW_PROP = "parser.headerRow";
    /**
     * The name of the property which specified the column delimiter for the
     * property {@link #COLUMN_NAMES_PROP}. Default ,
     */
    public static final String SECTION_FIELD_NAME_PROP = "parser.section.field.name";
    /**
     * The name of the property which specified the column delimiter for the
     * property {@link #COLUMN_NAMES_PROP}. Default ,
     */
    public static final String COLUMN_DELIM_PROP = "parser.columns.delim";
    /**
     * The name of the property which specifies the Class name that builds the
     * Hierarchy of Contents. Since Excel is only a flat structure it may be
     * necessary to specify how to structure the Content inside Integrity. Since
     * the methodology to do this is unique to the situation, this provides a
     * way to do this. Your class must extend
     * {@link com.integrity.gateway.parser.positioner.AbstractPositioner#AbstractPositioner(List, Properties)}.
     */
    // VE: public static final String BUILD_HIERARCHY_PROP = "parser.hierarchyClass";
    /**
     * The name of the property which specifies if blank rows should be skipped.
     * Default true.
     */
    public static final String IGNORE_BLANKS_PROP = "parser.ignoreBlankRows";
    /**
     * Allows you to transform/combine column values. In combinations with
     * {@link #TRANSFORM_FIELDS_NAME_SUFFIX} and
     * {@link #TRANSFORM_FIELDS_PATTERN_SUFFIX}, The name of the propert[y/ies]
     * that specify how to transform the columns into what the Gateway Mapper
     * should be sent.<br>
     * The properties should be of form:
     * <br><pre>
     *  parser.transformColumns.#.Name=...
     *  parser.transformColumns.#.Pattern=...
     * </pre>
     * <br>
     * The # must be sequential starting from '1', order is not important and
     * you can have as many as you require. Also there must be one property each
     * for the Name and Pattern suffixes.
     *
     * @see #TRANSFORM_FIELDS_NAME_SUFFIX
     * @see #TRANSFORM_FIELDS_PATTERN_SUFFIX
     *
     */
    public static final String TRANSFORM_FIELDS_PREFIX_PROP = "parser.transformColumns";

    /**
     * The property name that specified the first row to read in.
     */
    public static final String START_ROW_PROP = "parser.startRow";

    /**
     * The property name that specified the last row to read in.
     */
    public static final String END_ROW_PROP = "parser.endRow";
    /**
     * The Name of the transformed field which will be used in the Gateway
     * Mapper
     *
     * @see #TRANSFORM_FIELDS_PREFIX_PROP
     */
    public static final String TRANSFORM_FIELDS_NAME_SUFFIX = "parser.transformNameSuffix";
    /**
     * The Pattern of what the newly created field should look like.
     * <br>This is specified by any String where the values of other columns are
     * specified like:
     * <tt>{columnName}</tt>. If the column contents are empty then it is
     * replaced with an empty String.
     *
     * @see #TRANSFORM_FIELDS_PREFIX_PROP
     */
    public static final String TRANSFORM_FIELDS_PATTERN_SUFFIX = "parser.transformPatternSuffix";

    /**
     * This is the property name for the Document Title Field which will be used
     * by the mapper.
     */
    // VE: public static final String DOC_TITLE_FIELD_PROP = "gateway.docTitleField";
    /**
     * This is the property name for the Document Description Field which will
     * be used by the mapper.
     */
    // VE: public static final String DOC_DESC_FIELD_PROP = "gateway.docDescField";
    /**
     * Boolean property to set if this is a new document and the Gateway
     * interface will not require an ID.
     */
    // public static final String IS_NEW_DOC_PROP = "gateway.isNewDocument";
    /**
     * The link field from the map. This is required if {@link #IS_NEW_DOC_PROP}
     * is true.
     */
    public static final String DOCUMENT_LINK_FIELD_PROP = "gateway.documentLinkField";
    public static final String CONTENT_LINK_FIELD_PROP = "gateway.contentLinkField";

    public static final String RELATIONSHIP_FIELD_NAME_PROP = "parser.relationshipFieldName";

    /**
     * Should debug logging be presented (System.err and System.out)
     */
    public static boolean DEBUG;
    /**
     * This is a global counter to increment and uniquely identify each item
     */
    public static int itemCount = 1;
    // public static String documentID;

    public static List<ExternalItem> relationships;
    public static ItemMapperSession imSession;

    public void initParams() {
        log("------------------------------------------------------------------", 1);
        log("Listing of Parameters:", 1);
        pSheetNumber = getConfigParameter(SHEET_NUMBER_PROP, "1", "int");
        pStartRow = getConfigParameter(START_ROW_PROP, "1", "int");
        // no more used: will be determined automatically
        // pHeaderRow = getConfigParameter(HEADER_ROW_PROP, "6", "int");
        pSectionFieldName = getConfigParameter(SECTION_FIELD_NAME_PROP, "Section", "field");
        pEndRow = getConfigParameter(END_ROW_PROP, "0", "int");
        pDocumentLinkField = getConfigParameter(DOCUMENT_LINK_FIELD_PROP, "Document ID", "field");
        pContentLinkField = getConfigParameter(CONTENT_LINK_FIELD_PROP, "External ID", "field");
        pColumnDelim = getConfigParameter(COLUMN_DELIM_PROP, ",", "char");
        pColumnNames = getConfigParameter(COLUMN_NAMES_PROP, "", "String");
        pDebug = getConfigParameter(DEBUG_PROP, "true", "boolean");
        pIgnoreBlanks = getConfigParameter(IGNORE_BLANKS_PROP, "true", "boolean");
        pMapPosition = getConfigParameter(MAP_POSITION_PROP, "false", "boolean");
        pSectionColumnName = getConfigParameter(SECTION_COLUMN_NAME_PROP, "Section", "field");
        // pPropertiesFilePath = getConfigParameter(PROPERTIES_FILE_PATH, "n/a", "String");
        pRelationshipFieldName = getConfigParameter(RELATIONSHIP_FIELD_NAME_PROP, "", "field");
        pTransformFieldsNameSuffix = getConfigParameter(TRANSFORM_FIELDS_NAME_SUFFIX, "", "String");
        pTransformFieldsPatternSuffix = getConfigParameter(TRANSFORM_FIELDS_PATTERN_SUFFIX, "", "String");
        pTransformFieldsPrefix = getConfigParameter(TRANSFORM_FIELDS_PREFIX_PROP, "", "String");
        log("------------------------------------------------------------------", 1);
    }

    /**
     * Get Config Parameter from Gateway-Tool-Config.xml file
     *
     * @param param
     * @param defaultValue
     * @param type
     * @return
     */
    public String getConfigParameter(String param, String defaultValue, String type) {
        String configParam = getConfigParam(param);
        if (configParam != null && !configParam.contentEquals("")) {

            if (type.toLowerCase().contentEquals("int")) {
                try {
                    Integer.parseInt(type);
                } catch (NumberFormatException ex) {
                    exitProgram("ERROR: Property " + param + " isn't a number, please correct!");
                }
            }
            if (type.toLowerCase().contentEquals("boolean")) {
                try {
                    Boolean.parseBoolean(type);
                } catch (NumberFormatException ex) {
                    exitProgram("ERROR: Property " + param + " isn't a boolean, please correct!");
                }
            }
        }
        String finalValue = (configParam == null || configParam.contentEquals("") ? defaultValue : configParam);
        if (configParam == null || configParam.contentEquals("")) {
            log("Parameter " + param + " (" + type + ") is unset, taking default '" + finalValue + "'", 2);
        } else {
            log("Parameter " + param + " (" + type + ") is set to '" + finalValue + "'", 2);
        }
        return finalValue;
    }

    public static void exitProgram(String text) {
        log(text, 1);
        exit(-1);
    }

    /**
     * Builds a new IIF item.
     *
     * @param mappingConfig
     * @param content the Content to transform to IIF
     * @param relFieldName
     * @param linkContentField
     * @return the Item to be added to the Document.
     * @throws com.mks.gateway.mapper.UnsupportedPrototypeException
     */
    public static ExternalItem buildContentItem(ItemMapperConfig mappingConfig, Content content, String relFieldName, String linkContentField) throws UnsupportedPrototypeException, ItemMapperException, APIException {
        HashMap<String, String> fields = content.getFieldValuesMap();
        String lcf = fields.get(linkContentField);
        ExternalItem item = new ExternalItem(content.isDocument() ? "DOCUMENT" : "CONTENT", lcf);;

        if (content.isDocument()) {
            item.setInternalId(content.getFieldValue("ID").substring(content.getFieldValue("ID").indexOf("-") + 1));
        }

        // Add Default values defined in the XML Mapping file
        for (ItemMapperConfig.Field fieldConfig : mappingConfig.getIncomingFields()) {
            if (fieldConfig.defaultValue != null && !fieldConfig.defaultValue.isEmpty()) {
                item.add(fieldConfig.externalField, fieldConfig.defaultValue);
            }
        }

        // item.add(linkContentField, fields.get(fields.get(linkContentField)));
        // VE: item.setSourceSystemID("ITEM" + itemCount++);
        // item.setSourceSystemID(fields.get(linkContentField));
        for (String field : fields.keySet()) {
            if (MappingConfig.containsExternalField(mappingConfig, field)) {

                // log("INFO: Field Type for " + field + " is: " + MappingConfig.getExternalFieldType(mappingConfig, field), 1);
                if (MappingConfig.getExternalFieldType(mappingConfig, field).contentEquals("richcontent")) {
                    // item.add(field, fields.get(field));

                    String richcontent = fields.get(field);
                    // not here!
                    // richcontent = richcontent.replace("<", "&lt;").replace(">", "&gt;");

                    // replace the charage returns
//                    if (richcontent.contains("" + (char) (10))) {
//                        // log("Char 10 found, len=" + text.length(), 2);
//                        richcontent = richcontent.replace("" + (char) (10), "<br>");
//                    }
                    ItemField itemField = new com.mks.gateway.data.ItemField.RichContent(field, "<!-- MKS HTML -->" + richcontent);
                    item.getItemData().addField(itemField);

                } else if (MappingConfig.getExternalFieldType(mappingConfig, field).contentEquals("relationship")) {
                    // log("field '" + field + "': " + fields.get(field), 2);
                    // if (field.contentEquals("External ID")) {

                    if (!ExcelParser.buildRelationship.contentEquals("false")) {
                        // if (!relFieldName.isEmpty()) {
                        String itemID = fields.get(field).replaceAll("[^0-9.,]", "");
                        addRelationship(item, field, itemID);
                        // }
                    }
                    // }
                } else {

                    ItemField itemField = new com.mks.gateway.data.ItemField(field, fields.get(field));
                    item.getItemData().addField(itemField);
                }
            }
        }
        if (content.hasRelationships() && relFieldName != null) {
            if (relationships == null) {
                relationships = new ArrayList<ExternalItem>();
            }
            List<Content> rels = content.getRelationships();
            int relCount = 0;

//            Relationship relationship = new Relationship(relFieldName);
//            for (Content rel : rels) {
//                String relID = "REL_ITEM" + (itemCount - 1) + "_" + ++relCount;
//                ExternalItem relItem = new ExternalItem("ISSUE", relID);
//                // relItem.setSourceSystemID(relID);
//                // relItem.setIifPrototype("ISSUE");
//                HashMap<String, String> relFields = rel.getFieldValuesMap();
//                for (String field : relFields.keySet()) {
//                    relItem.add(field, relFields.get(field));
//                }
//                relationships.add(relItem);
//                relationship.addTarget(new Target(relID));
//            }
//            item.addRelatedItem(relationship);
        }
        for (Content child : content.getChildren()) {
            if (child != null && child.getFieldValuesMap().size() > 0) {
                item.addChild(buildContentItem(mappingConfig, child, relFieldName, linkContentField));
            }
        }
        return item;
    }

    public static boolean setCurrentDirectory(String directory_name) {
        boolean dirResult = false;  // Boolean indicating whether directory was set
        File directory;       // Desired current working directory

        directory = new File(directory_name).getAbsoluteFile();
        if (directory.exists() || directory.mkdirs()) {
            dirResult = (setProperty("user.dir", directory.getAbsolutePath()) != null);
        }

        return dirResult;
    }

    /**
     * addRelationship
     *
     * @param item
     * @param relationshipName
     * @param relatedItemID
     * @throws UnsupportedPrototypeException
     * @throws ItemMapperException
     */
    public static void addRelationship(ExternalItem item, String relationshipName, String relatedItemID) throws UnsupportedPrototypeException, ItemMapperException {
        // relatedItem.getId(): Item.CONTENT [internalID:687, externalID:__MKSID__687]
        ExternalItem relItem = new ExternalItem("CONTENT", "__MKSID__" + relatedItemID);
        relItem.setInternalId(relatedItemID);

        // ExternalItem relatedItem = getNewItem(imSession, mappingConfig, ExcelParser.iGatewayDriver, itemID, null);
        // log ("relatedItem.getId(): " + relItem.getId(), 1);
        ItemField itemField = new com.mks.gateway.data.ItemField.Relationship(relationshipName, relItem.getId());
        item.getItemData().addField(itemField);

    }

    /**
     * This creates the MKSItems Element of the IIF to be exported. This is a
     * recursive algorithm which goes through each Content Object in
     * <tt>finalContent</tt> recursively adding all of its children.
     *
     * @param mappingConfig
     * @param finalContent the final Content List to be parsed. If Hierarchy is
     * important this is the List returned from the concrete Class that
     * implements
     * {@link com.integrity.gateway.parser.positioner.AbstractPositioner}
     * @param title the title of the document
     * @param defaultItemFields
     * @return the IIF as a MKSItems Objects which can then be serialized to
     * disk.
     * @throws com.mks.gateway.mapper.UnsupportedPrototypeException
     */
    public static List<ExternalItem> createIIFList(ItemMapperConfig mappingConfig, List<Content> finalContent, String title, Map<String, String> defaultItemFields) throws UnsupportedPrototypeException, ItemMapperException, APIException {

        ItemMapperConfig docMappingConfig = MappingConfig.getDocumentMappingConfig(mappingConfig);
        ItemMapperConfig contentMappingConfig = MappingConfig.getContentMappingConfig(mappingConfig);

        ExternalItem docItem = new ExternalItem("DOCUMENT", title);
        // doc.setSourceSystemID(title);

        // if any default value is given in the load screen, this will be taken at first
        for (String key : defaultItemFields.keySet()) {
            if (defaultItemFields.get(key) != null && !defaultItemFields.get(key).isEmpty()) {
                docItem.add(key, defaultItemFields.get(key));
            }
        }

        // now, adds all document fields
        for (Field field : docFields) {
            if (MappingConfig.containsExternalField(docMappingConfig, field.getFieldName())) {

                // if the field is already there, then the field will NOT overwrite whats there already
                if (!docItem.hasField(field.getFieldName())) {
                    docItem.getItemData().addField(field.getFieldName(), field.getFieldValue());
                }
            }
        }

        // 20 Levels shall be enought
        ExternalItem[] sectionList = new ExternalItem[20];

        // Rebuild the hierarchy
        for (Content content : finalContent) {
            if (content != null && content.getFieldValuesMap().size() > 0) {
                ExternalItem newItem = buildContentItem(contentMappingConfig, content,
                        pRelationshipFieldName,
                        pContentLinkField);
                int sectionLevel = content.getSectionLevel();

                // debug("SectioN Level " + sectionLevel + ", " + content.getPosition());
                if (sectionLevel > 1) {
                    sectionList[sectionLevel - 1].addChild(newItem);
                } else {
                    docItem.addChild(newItem);
                }
                // remember the new "last" item in this level
                sectionList[sectionLevel] = newItem;
            }
        }

        if (relationships != null) {
            for (ExternalItem relationshipItem : relationships) {
                docItem.addChild(relationshipItem);
            }
        }

        // Build the final itemObject
        List<ExternalItem> items = new ArrayList<ExternalItem>(1);
        items.add(docItem);
        return items;
    }

    /**
     * Prints out the List of Content
     *
     * @param contents the list of Content
     */
    public static void print(List<Content> contents) {
        for (Content c : contents) {
            if (c != null) {
                GatewayLogger.log(c.getPosition());
                print(c.getChildren());
            }
        }
    }

//    public static MappingItem getNewItem(
//            ItemMapperSession imSession,
//            ItemMapperConfig mappingConfig,
//            IGatewayDriver gatewayDriver,
//            String itemID,
//            Date asOf
//    ) throws GatewayException, ItemMapperException, APIException {
//
//        MappingItem item = MappingItem.newContentItem(Integer.parseInt(itemID), asOf);
//        item.setItemAttribute(ExternalItem.Attribute.ADAPTER.getStringValue(), "com.mks.gateway.mapper.bridge.IMAdapter");
//        IMAdapter adapter = (IMAdapter) IMAdapter.getAdapter(imSession, mappingConfig, item, gatewayDriver);
//        try {
//            adapter.retrieveItem(item, mappingConfig, false, asOf);
//        } catch (InvalidCommandOptionException ex) {
//            adapter.retrieveItem(item, mappingConfig, false, null);
//        }
//
//        Transformer.transform(item, imSession, mappingConfig, false, true, asOf);
//        return item;
//    }
}
