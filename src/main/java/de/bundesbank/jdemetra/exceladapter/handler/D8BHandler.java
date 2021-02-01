/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import de.bbk.concur.FixedOutlier;
import de.bbk.concur.util.D8BInfos;
import de.bundesbank.jdemetra.exceladapter.Data;
import de.bundesbank.jdemetra.exceladapter.ExcelMetaDataHelper;
import ec.nbdemetra.ui.properties.l2fprod.ColorChooser;
import ec.tss.Ts;
import ec.tss.sa.SaItem;
import ec.tstoolkit.timeseries.regression.OutlierEstimation;
import ec.tstoolkit.timeseries.simplets.TsData;
import ec.tstoolkit.timeseries.simplets.TsDataTable;
import ec.tstoolkit.timeseries.simplets.TsDataTableInfo;
import ec.tstoolkit.timeseries.simplets.TsDomain;
import ec.tstoolkit.timeseries.simplets.TsPeriod;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = AbstractHandler.class)
public class D8BHandler extends AbstractHandler {

    private static final String SUFFIX = "-D8B";
    private static final String NAME = "D8B",
            OPTION_ID = "exceladapter.d8b";
    private static final boolean DEFAULT = false;

    @Override
    protected List<Ts> extractData(List<SaItem> items) {
        return new ArrayList<>();
    }

    @Override
    public boolean writeWorksheet(Data data, XSSFWorkbook workbook, CellStyle dateStyle) {
        List<D8BInfos> infos = data.getCurrent().stream().map(SaItem::toDocument).map(D8BInfos::new).collect(Collectors.toList());

        List<String> names = data.getCurrent().stream().map(ExcelMetaDataHelper::getItemName).collect(Collectors.toList());
        TsDataTable dataTable = new TsDataTable();

        infos.forEach((info) -> {
            dataTable.add(info.getSi().getTsData());
        });

        TsDomain dom = dataTable.getDomain();
        if (dom == null || dom.isEmpty()) {
            return false;
        }

        XSSFSheet sheet = workbook.createSheet(data.getName() + getSuffix());
        Row header = sheet.createRow(0);
        int columnCount = 0;
        for (String name : names) {
            ++columnCount;
            Cell cell = header.createCell(columnCount);
            cell.setCellValue(name);
        }

        // write each row
        for (int j = 0; j < dom.getLength(); ++j) {
            Row row = sheet.createRow(j + 1);
            Cell dateCell = row.createCell(0);
            TsPeriod pointInTime = dom.get(j);
            dateCell.setCellValue(pointInTime.firstday().getTime());
            dateCell.setCellStyle(dateStyle);
            for (int i = 0; i < dataTable.getSeriesCount(); ++i) {
                TsDataTableInfo dataInfo = dataTable.getDataInfo(j, i);
                if (dataInfo == TsDataTableInfo.Valid) {
                    Cell cell = row.createCell(i + 1);
                    cell.setCellValue(dataTable.getData(j, i));
                    D8BInfos d8BInfos = infos.get(i);
                    useAdditionalInfo(workbook, sheet, cell, d8BInfos, pointInTime);
                }
            }
        }
        sheet.createFreezePane(1, 1);
        sheet.setColumnWidth(0, 10 * 256);
        return true;
    }


    @Override
    public String getSuffix() {
        return SUFFIX;
    }

    @Override
    public String getName() {
        return NAME;
    }

    @Override
    public String getOptionID() {
        return OPTION_ID;
    }

    @Override
    public boolean getDefault() {
        return DEFAULT;
    }

    private void useAdditionalInfo(XSSFWorkbook workbook, XSSFSheet sheet, Cell cell, D8BInfos x, TsPeriod period) {
        StringBuilder commentBuilder = new StringBuilder();
        OutlierEstimation[] outliers = x.getBoth();
        FixedOutlier[] fixedOutliers = x.getFixedOutliers();
        TsData d9 = x.getReplacementValues();


        boolean found = false;
        for (OutlierEstimation outlier : outliers) {
            if (outlier == null) {
                continue;
            }
            if (outlier.getPosition().equals(period)) {
                commentBuilder.append("Outlier Value : ")
                        .append(format(outlier.getValue())).append("\n")
                        .append("TStat : ")
                        .append(format(outlier.getTStat())).append("\n")
                        .append("Outlier type : ")
                        .append(outlier.getCode()).append("\n");
                XSSFCellStyle style = workbook.createCellStyle();
                style.setFillForegroundColor(new XSSFColor(ColorChooser.getColor(outlier.getCode()), null));
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cell.setCellStyle(style);
                found = true;
                break;
            }
        }

        if (!found) {
            for (FixedOutlier fixedOutlier : fixedOutliers) {
                if (fixedOutlier == null) {
                    continue;
                }
                if (fixedOutlier.getPosition().equals(period)) {
                    commentBuilder.append("Fixed Outlier Value : ")
                            .append(format(fixedOutlier.getValue())).append("\n")
                            .append("Outlier type : ")
                            .append(fixedOutlier.getCode()).append("\n");
                    XSSFCellStyle style = workbook.createCellStyle();
                    style.setFillForegroundColor(new XSSFColor(ColorChooser.getColor(fixedOutlier.getCode()), null));
                    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cell.setCellStyle(style);
                    break;
                }
            }

        }

        if (d9 != null && d9.getFrequency() == period.getFrequency() && Double.isFinite(d9.get(period))) {
            commentBuilder.append("Extreme Value Replacement : ")
                    .append(format(d9.get(period)));

        }
        String commentText = commentBuilder.toString();
        if (!commentText.isEmpty()) {

            CreationHelper factory = workbook.getCreationHelper();
            Drawing drawing = sheet.createDrawingPatriarch();

            ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex());
            anchor.setCol2(cell.getColumnIndex() + 4);
            anchor.setRow1(cell.getRowIndex());
            anchor.setRow2(cell.getRowIndex() + 5);

            Comment comment = drawing.createCellComment(anchor);
            RichTextString str = factory.createRichTextString(commentText);
            comment.setString(str);
            comment.setAuthor("JDemetra");

            cell.setCellComment(comment);
        }
    }

    private String format(double value) {
        return String.format("%.3f", value);
    }
}
