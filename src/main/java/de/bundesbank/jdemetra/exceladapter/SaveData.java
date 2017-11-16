/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import ec.satoolkit.x11.X11Kernel;
import ec.tss.Ts;
import ec.tss.TsFactory;
import ec.tss.sa.SaItem;
import ec.tstoolkit.algorithm.CompositeResults;
import ec.tstoolkit.timeseries.simplets.TsData;
import ec.tstoolkit.timeseries.simplets.TsDataTable;
import ec.tstoolkit.timeseries.simplets.TsDataTableInfo;
import ec.tstoolkit.timeseries.simplets.TsDomain;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Thomas Witthohn
 */
public class SaveData {

    private Map<String, List<Ts>> infos = new HashMap<>();

    public void addData(String multiDocName, List<SaItem> items) {
        infos.put(multiDocName, extractInfos(items));
    }

    public boolean save() {
        File myFile = new File("C://Daten/test.xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook();

        for (Map.Entry<String, List<Ts>> entry : infos.entrySet()) {
            XSSFSheet sheet = workbook.createSheet(entry.getKey());

            List<Ts> value = entry.getValue();
            TsDataTable dataTable = new TsDataTable();

            Row row = sheet.createRow(0);
            int columnCount = 0;
            for (Ts ts : value) {
                dataTable.insert(-1, ts.getTsData());
                Cell cell = row.createCell(++columnCount);
                cell.setCellValue(ts.getRawName());
            }

            TsDomain dom = dataTable.getDomain();
            if (dom == null || dom.isEmpty()) {
                //TODO?
                continue;
            }
            // write each rows
            for (int j = 0; j < dom.getLength(); ++j) {
                Row row2 = sheet.createRow(j + 1);
                row2.createCell(0).setCellValue(dom.get(j).toString());
                for (int i = 0; i < dataTable.getSeriesCount(); ++i) {
                    TsDataTableInfo dataInfo = dataTable.getDataInfo(j, i);
                    switch (dataInfo) {
                        case Valid:
                            Cell cell = row2.createCell(i + 1);
                            cell.setCellValue(dataTable.getData(j, i));
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        try (FileOutputStream outputStream = new FileOutputStream("C:\\Daten\\Testssssss.xlsx")) {
            workbook.write(outputStream);
        } catch (IOException ex) {
            Logger.getLogger(SaveData.class.getName()).log(Level.SEVERE, null, ex);
        }

        return true;
    }

    private List<Ts> extractInfos(List<SaItem> item) {
        List<Ts> set = new ArrayList<>();
        item.stream().map((saItem) -> {
            String name = saItem.getRawName().isEmpty() ? saItem.getTs().getRawName() : saItem.getRawName();
            CompositeResults compositeResults = saItem.process();
            TsData d10Data = compositeResults.getData(X11Kernel.D10, TsData.class);
            Ts d10 = TsFactory.instance.createTs(name + ".d10", null, d10Data);
            return d10;
        }).forEachOrdered((d10) -> {
            set.add(d10);
        });
        return set;
    }

}
