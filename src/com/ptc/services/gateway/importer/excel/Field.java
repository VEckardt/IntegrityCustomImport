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

/**
 * "Simle" object that represents the Field IIF Tag
 *
 * @author veckardt
 */
public class Field {

    public Field(String fieldName, String fieldValue) {
        this.fieldName = fieldName;
        this.fieldValue = fieldValue;
    }

    public Field(String fieldName, String fieldValue, String fieldType) {
        this.fieldName = fieldName;
        this.fieldValue = fieldValue;
        this.fieldType = fieldType;
    }

    public Field(String fieldName) {
        this.fieldName = fieldName;
        this.fieldValue = "";
    }

    private String fieldValue = "";

    private String fieldName = "";
    //@Attribute(name = "data-type")

    private String fieldType = "string";

    public void setFieldType(String dataType) {
        this.fieldType = dataType;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public String getFieldValue() {
        return fieldValue;
    }

    public String getFieldType() {
        return fieldType;
    }

    public Boolean isRelationship() {
        return fieldType.contentEquals("relationship");
    }

    public void setFieldValue(String fieldValue) {
        this.fieldValue = fieldValue;
    }

}
