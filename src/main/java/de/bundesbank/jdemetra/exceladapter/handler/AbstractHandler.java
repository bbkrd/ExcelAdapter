/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import de.bundesbank.jdemetra.exceladapter.Data;
import de.bundesbank.jdemetra.exceladapter.options.ExcelAdapterOptionsPanelController;
import ec.tss.Ts;
import ec.tss.sa.SaItem;
import ec.tstoolkit.design.ServiceDefinition;
import ec.tstoolkit.timeseries.simplets.TsDataTable;
import ec.tstoolkit.timeseries.simplets.TsDataTableInfo;
import ec.tstoolkit.timeseries.simplets.TsDomain;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openide.util.NbPreferences;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceDefinition
public abstract class AbstractHandler {

    protected abstract List<Ts> extractData(List<SaItem> items);

    public boolean writeWorksheet(Data data, XSSFWorkbook workbook, CellStyle dateStyle) {
        List<Ts> value = extractData(data.getCurrent());
        TsDataTable dataTable = new TsDataTable();
        value.forEach((ts) -> {
            dataTable.add(ts.getTsData());
        });

        TsDomain dom = dataTable.getDomain();
        if (dom == null || dom.isEmpty()) {
            return false;
        }

        XSSFSheet sheet = workbook.createSheet(data.getName() + getSuffix());
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
            dateCell.setCellStyle(dateStyle);
            for (int i = 0; i < dataTable.getSeriesCount(); ++i) {
                TsDataTableInfo dataInfo = dataTable.getDataInfo(j, i);
                if (dataInfo == TsDataTableInfo.Valid) {
                    Cell cell = row.createCell(i + 1);
                    cell.setCellValue(dataTable.getData(j, i));
                }
            }
        }
        sheet.createFreezePane(1, 1);
        sheet.setColumnWidth(0, 10 * 256);
        return true;
    }

    public boolean isEnabled() {
        return NbPreferences.forModule(ExcelAdapterOptionsPanelController.class).getBoolean(getOptionID(), getDefault());
    }

    public abstract String getSuffix();

    public abstract String getName();

    public abstract String getOptionID();

    public abstract boolean getDefault();

    @Override
    public String toString() {
        return getName();
    }

}
