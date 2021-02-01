/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import de.bbk.concur.util.SavedTables;
import static de.bundesbank.jdemetra.exceladapter.ExcelConnection.*;
import de.bundesbank.jdemetra.exceladapter.handler.A1AHandler;
import de.bundesbank.jdemetra.exceladapter.handler.A6Handler;
import de.bundesbank.jdemetra.exceladapter.handler.D10Handler;
import ec.tss.sa.SaItem;
import ec.tstoolkit.MetaData;
import java.util.List;
import java.util.function.BiFunction;

/**
 *
 * @author Deutsche Bundesbank
 */
public class ExcelMetaDataHelper {

    private final String filePath;

    public ExcelMetaDataHelper() {
        this.filePath = "";
    }

    public ExcelMetaDataHelper(String filePath) {
        if (filePath == null) {
            this.filePath = "";
        } else {
            this.filePath = filePath;
        }
    }

    public static String getItemName(SaItem item) {
        return item.getRawName().isEmpty() ? item.getTs().getRawName() : item.getRawName();
    }

    public void addMetaData(Data item) {
        String multiDocName = item.getName();
        for (SaItem saItem : item.getCurrent()) {
            String itemName = getItemName(saItem);
            MetaData metaData = getOrCreateMetaData(saItem);
            fillMetaData(metaData::putIfAbsent, multiDocName, itemName);
        }
    }

    public void addMetaData(List<Data> items) {
        items.forEach(this::addMetaData);
    }

    public void overrideMetaData(Data item) {
        String multiDocName = item.getName();
        for (SaItem saItem : item.getCurrent()) {
            String itemName = getItemName(saItem);
            MetaData metaData = getOrCreateMetaData(saItem);
            fillMetaData(metaData::put, multiDocName, itemName);
        }
    }

    public void overrideMetaData(List<Data> items) {
        items.forEach(this::overrideMetaData);
    }

    private void fillMetaData(BiFunction<String, String, String> function, String multiDocName, String itemName) {
        for (SavedTables.TABLES table : SavedTables.TABLES.values()) {
            String tableName;
            switch (table) {
                case CALENDARFACTOR:
                    tableName = A6Handler.SUFFIX;
                    break;
                case FORECAST:
                    tableName = A1AHandler.SUFFIX;
                    break;
                case SEASONALFACTOR:
                    tableName = D10Handler.SUFFIX;
                    break;
                default:
                    tableName = "";
            }
            function.apply(METADATA_EXCEL_FILE + table, filePath);
            function.apply(METADATA_EXCEL_SHEET + table, multiDocName + tableName);
            function.apply(METADATA_EXCEL_SERIES + table, itemName);
        }
    }

    private MetaData getOrCreateMetaData(SaItem item) {
        MetaData metaData = item.getMetaData();
        if (metaData == null) {
            metaData = new MetaData();
            item.setMetaData(metaData);
        }
        return metaData;
    }
}
