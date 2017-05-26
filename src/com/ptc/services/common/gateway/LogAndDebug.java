/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ptc.services.common.gateway;

import com.mks.gateway.GatewayLogger;
import com.mks.gateway.data.ExternalItem;
import com.mks.gateway.data.xml.ItemXMLConstants;
import com.mks.gateway.data.xml.ItemXMLWriter;
import com.mks.gateway.mapper.ItemMapperException;
import java.io.File;
import java.util.Iterator;
import java.util.List;

/**
 *
 * @author veckardt
 */
public class LogAndDebug {

    Boolean debug = true;

    public LogAndDebug(Boolean debug) {
        this.debug = debug;
    }

    /**
     *
     * @param string
     * @param level
     */
    public static void log(String string, int level) {
        String str = "          ".substring(0, level - 1);
        if (GatewayLogger.getLogFile() == null) {
            System.out.println("(" + level + ") " + str + string);
        } else {
            GatewayLogger.logMessage(str + string, level);
        }
    }

    /**
     *
     * @param level
     */
    public static void logLine(int level) {
        String string = "--------------------------------------------------------------------------";
        log(string, 1);
    }

    /**
     * writes the given list of ExternalItem to file. file will be created in
     * the temp directory of the jvm
     *
     * @param gatewayItems
     */
    public void writeIIFtoDisk(List<ExternalItem> gatewayItems, int id) {
        String path = "C:\\IntegrityWordExport\\";
        if (debug) {
            File f = new File(path);
            if (f.exists() && f.isDirectory()) {

                try {
                    // File out = File.createTempFile("c:\\temp\\WORDEXPORT_", ".iif");
                    File out = new File(path + "WordExport_iif_" + Integer.toString(id) + ".xml");
                    ItemXMLWriter xmlWriter = new ItemXMLWriter(out, ItemXMLConstants.SCHEMA_VERSION_1_0);
                    xmlWriter.write(gatewayItems);
                    GatewayLogger.logMessage("DEBUG: Successfully created debug iif-file " + out.getAbsolutePath(), 5);
                } catch (Exception e) {
                    GatewayLogger.logError(e, 10);
                }
            }
        }
    }

    public void writeIIFtoDisk(List<ExternalItem> gatewayItems) {
        try {
            File out = File.createTempFile("WORDEXPORT_", ".iif");
            ItemXMLWriter xmlWriter = new ItemXMLWriter(out, ItemXMLConstants.SCHEMA_VERSION_1_0);
            xmlWriter.write(gatewayItems);
            GatewayLogger.logMessage("Successfully created debug iif-file " + out.getAbsolutePath(), 5);
        } catch (Exception e) {
            GatewayLogger.logError(e, 10);
        }
    }

    /**
     * recurses down over all child elements until no more childs are available
     * and logs the fields..
     *
     * @param ei starting ExternalItem
     * @param result
     */
    public void recurseChilds(ExternalItem ei, List<String> result) {
        try {
            if (!ei.getChildren().isEmpty()) {
                //there are childs attached..
                Iterator<ExternalItem> childs = ei.childrenIterator();
                while (childs.hasNext()) {
                    ExternalItem currentChild = childs.next();
                    GatewayLogger.logMessage("child: " + currentChild.getField("Section").getStringValue() + ": " + currentChild.toString(), 1);
                    GatewayLogger.logMessage("Available Fields:", 1);
                    Iterator<?> availableFields = currentChild.getFieldNames().iterator();
                    while (availableFields.hasNext()) {
                        GatewayLogger.logMessage("Field: " + availableFields.next().toString(), 1);
                    }
                    recurseChilds(currentChild, result);
                }
            }
        } catch (ItemMapperException e) {
            GatewayLogger.logError(e.getMessage(), e, 10);
            if (result != null) {
                result.add("Problem while iterating over all childs: " + e.getMessage());
            }
        }
    }
}
