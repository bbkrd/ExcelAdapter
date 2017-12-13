/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import ec.satoolkit.DecompositionMode;
import ec.satoolkit.x11.X11Kernel;
import ec.tss.Ts;
import ec.tss.TsFactory;
import ec.tss.sa.SaItem;
import ec.tstoolkit.algorithm.CompositeResults;
import ec.tstoolkit.modelling.ModellingDictionary;
import ec.tstoolkit.timeseries.simplets.TsData;
import ec.tstoolkit.timeseries.simplets.TsDataTable;
import ec.tstoolkit.timeseries.simplets.TsDataTableInfo;
import ec.tstoolkit.timeseries.simplets.TsDomain;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Thomas Witthohn
 */
public class SaveData {

    private final Map<String, List<Ts>> infos = new TreeMap<>();
    private static final short DATE_FORMAT = 14;

    public void addData(String multiDocName, List<SaItem> items) {
        infos.put(multiDocName, extractInfos(items));
    }

    public boolean save() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(DATE_FORMAT);

        for (Map.Entry<String, List<Ts>> entry : infos.entrySet()) {
            List<Ts> value = entry.getValue();
            TsDataTable dataTable = new TsDataTable();
            value.forEach((ts) -> {
                dataTable.add(ts.getTsData());
            });

            TsDomain dom = dataTable.getDomain();
            if (dom == null || dom.isEmpty()) {
                //TODO?
                Logger.getLogger(SaveData.class.getName()).log(Level.INFO, "Nothing in {0}", entry.getKey());
                continue;
            }

            XSSFSheet sheet = workbook.createSheet(entry.getKey());
            Row header = sheet.createRow(0);
            int columnCount = 0;
            for (Ts ts : value) {
                ++columnCount;
                Cell cell = header.createCell(columnCount);
                cell.setCellValue(ts.getRawName());
            }

            // write each row
            for (int j = 0; j < dom.getLength(); ++j) {
                Row row = sheet.createRow(j + 1);
                Cell dateCell = row.createCell(0);
                dateCell.setCellValue(dom.get(j).firstday().getTime());
                dateCell.setCellStyle(cellStyle);
                for (int i = 0; i < dataTable.getSeriesCount(); ++i) {
                    TsDataTableInfo dataInfo = dataTable.getDataInfo(j, i);
                    if (dataInfo == TsDataTableInfo.Valid) {
                        Cell cell = row.createCell(i + 1);
                        cell.setCellValue(dataTable.getData(j, i));
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
        item.forEach((saItem) -> {
            String name = saItem.getRawName().isEmpty() ? saItem.getTs().getRawName() : saItem.getRawName();
            CompositeResults result = saItem.process();
            boolean isMultiplicative = result.getData(ModellingDictionary.MODE, DecompositionMode.class).isMultiplicative();

            TsData d10 = result.getData(X11Kernel.D10, TsData.class);
            if (d10 != null) {
                TsData d10a = result.getData(X11Kernel.D10a, TsData.class);
                Ts seasonalfactor = TsFactory.instance.createTs(name + ".seasonalfactor", null, d10.update(d10a));
                set.add(seasonalfactor);
            }

            TsData calFactorData;
            TsData a6 = result.getData(X11Kernel.A6, TsData.class);
            TsData a7 = result.getData(X11Kernel.A7, TsData.class);
            if (a6 == null) {
                calFactorData = a7;
            } else {
                if (isMultiplicative) {
                    calFactorData = a6.times(a7);
                } else {
                    calFactorData = a6.plus(a7);
                }
            }
            if (calFactorData != null) {
                Ts calFactor = TsFactory.instance.createTs(name + ".calendarfactor", null, calFactorData);
                set.add(calFactor);
            }

            TsData a1 = result.getData(X11Kernel.A1, TsData.class);
            if (a1 != null) {
                TsData a1a = result.getData(X11Kernel.A1a, TsData.class);
                Ts forecast = TsFactory.instance.createTs(name + ".forecast", null, a1.update(a1a));
                set.add(forecast);
            }
        });
        return set;
    }

}
