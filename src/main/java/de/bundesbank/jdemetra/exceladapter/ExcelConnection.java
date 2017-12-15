/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import de.bbk.concur.servicedefinition.IExternalDataProvider;
import ec.tss.Ts;
import ec.tss.TsFactory;
import ec.tss.TsInformationType;
import ec.tss.TsMoniker;
import ec.tss.tsproviders.DataSet;
import ec.tss.tsproviders.DataSource;
import ec.tss.tsproviders.spreadsheet.SpreadSheetBean;
import ec.tss.tsproviders.spreadsheet.SpreadSheetProvider;
import ec.tstoolkit.MetaData;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.openide.util.Lookup;
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = IExternalDataProvider.class)
public class ExcelConnection implements IExternalDataProvider {

    public static final String METADATA_EXCEL_FILE = "excel.file.";
    public static final String METADATA_EXCEL_SHEET = "excel.sheet.";
    public static final String METADATA_EXCEL_SERIES = "excel.series.";

    @Override
    public Ts convertMetaDataToTs(MetaData meta, String tableName) {
        String filePath = meta.get(METADATA_EXCEL_FILE + tableName);
        String sheetName = meta.get(METADATA_EXCEL_SHEET + tableName);
        String seriesName = meta.get(METADATA_EXCEL_SERIES + tableName);
        if (filePath == null || sheetName == null || seriesName == null) {
            return createInvalidTs(tableName, "Not all necessary information");
        }

        try {
            File file = new File(filePath);

            SpreadSheetProvider provider = Lookup.getDefault().lookup(SpreadSheetProvider.class);
            if (!provider.accept(file)) {
                return createInvalidTs(tableName, "File not supported");
            }

            SpreadSheetBean bean = new SpreadSheetBean();
            bean.setFile(file);
            DataSource spreadsheet = bean.toDataSource(SpreadSheetProvider.SOURCE, SpreadSheetProvider.VERSION);

            List<DataSet> sheets = provider.children(spreadsheet);
            Optional<DataSet> specifiedSheet = sheets.stream().filter(i -> i.getParam("sheetName").get().equals(sheetName)).findFirst();
            if (!specifiedSheet.isPresent()) {
                return createInvalidTs(tableName, "Sheet " + sheetName + " doesn't exist.");
            }

            List<DataSet> series = provider.children(specifiedSheet.get());
            Optional<DataSet> specifiedSeries = series.stream().filter(i -> i.getParam("seriesName").get().equals(seriesName)).findFirst();
            if (!specifiedSeries.isPresent()) {
                return createInvalidTs(tableName, "Series " + seriesName + " doesn't exist in " + sheetName + ".");
            }

            TsMoniker moniker = provider.toMoniker(specifiedSeries.get());
            return TsFactory.instance.createTs(tableName, moniker, TsInformationType.All);
        } catch (IllegalArgumentException | IOException ex) {
            Logger.getLogger(ExcelConnection.class.getName()).log(Level.SEVERE, null, ex);
            return createInvalidTs(tableName, "Critical error!");
        }
    }

    private Ts createInvalidTs(String name, String invalidDataCause) {
        Ts ts = TsFactory.instance.createTs(name);
        ts.setInvalidDataCause(invalidDataCause);
        return ts;
    }

}
