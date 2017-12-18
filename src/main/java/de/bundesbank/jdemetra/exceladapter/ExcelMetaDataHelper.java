/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import de.bbk.concur.util.SavedTables;
import static de.bundesbank.jdemetra.exceladapter.ExcelConnection.*;
import ec.tss.sa.SaItem;
import ec.tstoolkit.MetaData;
import java.util.function.BiFunction;

/**
 *
 * @author Deutsche Bundesbank
 */
public class ExcelMetaDataHelper {

    private final String filePath;
    private String itemName, multiDocName;
    private MetaData metaData;

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

    public void addMetaData(SaItem item, String multiDocName) {
        getItemName(item);
        getOrCreateMetaData(item);
        this.multiDocName = multiDocName;
        fillMetaData(metaData::putIfAbsent);
    }

    public void overrideMetaData(SaItem item, String multiDocName) {
        getItemName(item);
        getOrCreateMetaData(item);
        this.multiDocName = multiDocName;
        fillMetaData(metaData::put);
    }

    private void fillMetaData(BiFunction<String, String, String> function) {
        for (SavedTables.TABLES table : SavedTables.TABLES.values()) {
            String tableName;
            switch (table) {
                case CALENDARFACTOR:
                    tableName = "-A6A7";
                    break;
                case FORECAST:
                    tableName = "-A1a";
                    break;
                case SEASONALFACTOR:
                    tableName = "-D10";
                    break;
                default:
                    tableName = "";
            }
            function.apply(METADATA_EXCEL_FILE + table, filePath);
            function.apply(METADATA_EXCEL_SHEET + table, multiDocName + tableName);
            function.apply(METADATA_EXCEL_SERIES + table, itemName);
        }
    }

    private void getItemName(SaItem item) {
        this.itemName = item.getRawName().isEmpty() ? item.getTs().getRawName() : item.getRawName();
    }

    private void getOrCreateMetaData(SaItem item) {
        metaData = item.getMetaData();
        if (metaData == null) {
            metaData = new MetaData();
            item.setMetaData(metaData);
        }
    }
}
