/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ptc.services.gateway.importer;

import com.mks.gateway.mapper.ItemMapperException;
import com.mks.gateway.mapper.config.ItemMapperConfig;
import com.mks.gateway.tool.exception.GatewayException;
import static com.ptc.services.gateway.importer.LogAndDebug.log;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author veckardt
 */
public class MappingConfig {

    public static String mappingConfigDocument = "DOCUMENT";
    public static String mappingConfigContent = "CONTENT";

    /**
     *
     * @param mappingConfig
     * @param description
     * @throws com.mks.gateway.tool.exception.GatewayException
     */
    public static void listMappingConfig(ItemMapperConfig mappingConfig, String description) throws GatewayException {
        try {
            log("Mapping Config ID: " + mappingConfig.getId(), 2);
            Iterator<?> it = mappingConfig.getOutgoingFields().listIterator();
            while (it.hasNext()) {
                ItemMapperConfig.Field fld = (ItemMapperConfig.Field) it.next();
                log(" Mapping Field: " + fld.internalField + " => " + fld.externalField + " (" + fld.dataType + ", " + fld.fieldType + ", " + fld.attachmentField + ")", 2);
            }
            log("End of listing " + description + " configuration.", 2);

            for (ItemMapperConfig subMappingConfig : mappingConfig.getSubConfigs()) {
                listMappingConfig(subMappingConfig, subMappingConfig.getId());
            }
        } catch (ItemMapperException ex) {
            Logger.getLogger(MappingConfig.class.getName()).log(Level.SEVERE, null, ex);
            throw new GatewayException(ex);
        }
    }

    public static ItemMapperConfig getDocumentMappingConfig(ItemMapperConfig mappingConfig) {
        return mappingConfig.getSubConfig(MappingConfig.mappingConfigDocument);
    }

    public static ItemMapperConfig getContentMappingConfig(ItemMapperConfig mappingConfig) {
        return mappingConfig.getSubConfig(MappingConfig.mappingConfigContent);
    }

    /**
     *
     * @param mappingConfig
     * @param fieldName
     * @return
     * @throws ItemMapperException
     */
    public static String getExternalFieldType(ItemMapperConfig mappingConfig, String fieldName) throws ItemMapperException {
        // log("Mapping Config ID: " + mappingConfig.getId(), 2);
        Iterator<?> it = mappingConfig.getOutgoingFields().listIterator();
        while (it.hasNext()) {
            ItemMapperConfig.Field fld = (ItemMapperConfig.Field) it.next();
            if (fieldName.contentEquals(fld.externalField)) {
                return fld.fieldType;
            }
        }
        return fieldName;
    }

    public static Boolean containsExternalField(ItemMapperConfig mappingConfig, String fieldName) throws ItemMapperException {
        // log("Mapping Config ID: " + mappingConfig.getId(), 2);
        Iterator<?> it = mappingConfig.getIncomingFields().listIterator();
        while (it.hasNext()) {
            ItemMapperConfig.Field fld = (ItemMapperConfig.Field) it.next();
            if (fieldName.contentEquals(fld.externalField)) {
                return true;
            }
        }
        return false;
    }

    public static String getInternalFieldName(ItemMapperConfig mappingConfig, String fieldName) throws ItemMapperException {
        // log("Mapping Config ID: " + mappingConfig.getId(), 2);
        Iterator<?> it = mappingConfig.getOutgoingFields().listIterator();
        while (it.hasNext()) {
            ItemMapperConfig.Field fld = (ItemMapperConfig.Field) it.next();
            if (fieldName.contentEquals(fld.externalField)) {
                return fld.internalField;
            }
        }
        return fieldName;
    }

    /**
     *
     * @param mappingConfig
     * @param fieldName
     * @return
     * @throws ItemMapperException
     */
    public static String getExternalFieldName(ItemMapperConfig mappingConfig, String fieldName) throws ItemMapperException {
        // log("Mapping Config ID: " + mappingConfig.getId(), 2);
        Iterator<?> it = mappingConfig.getOutgoingFields().listIterator();
        while (it.hasNext()) {
            ItemMapperConfig.Field fld = (ItemMapperConfig.Field) it.next();
            if (fieldName.contentEquals(fld.internalField)) {
                return fld.externalField;
            }
        }
        return fieldName;
    }

}
