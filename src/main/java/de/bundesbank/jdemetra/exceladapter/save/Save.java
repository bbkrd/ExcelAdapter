/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.save;

import de.bundesbank.jdemetra.exceladapter.Data;
import de.bundesbank.jdemetra.exceladapter.handler.AbstractHandler;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;
import java.util.List;
import java.util.stream.Collectors;
import javax.swing.JOptionPane;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openide.util.Lookup;

/**
 *
 * @author Thomas Witthohn
 */
@Slf4j
@lombok.experimental.UtilityClass
public class Save {

    private static final short DATE_FORMAT = 14;

    public static boolean save(List<Data> items, File file) {
        List<? extends AbstractHandler> enabledHandler = Lookup.getDefault().lookupAll(AbstractHandler.class)
                .stream()
                .filter(AbstractHandler::isEnabled)
                .collect(Collectors.toList());
        try {
            File tempFile = File.createTempFile("JDemetra+", "ExcelSave.xlsx");
            log.info("Writing to {}", tempFile.getAbsolutePath());
            try (FileOutputStream tempStream = new FileOutputStream(tempFile)) {
                XSSFWorkbook workbook = new XSSFWorkbook();
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setDataFormat(DATE_FORMAT);

                items.stream()
                        .filter(item -> !item.getCurrent().isEmpty())
                        .forEach(item -> enabledHandler.forEach(iHandler -> iHandler.writeWorksheet(item, workbook, cellStyle)));

                workbook.write(tempStream);

                boolean tryAgain;
                do {
                    try {
                        Files.copy(tempFile.toPath(), file.toPath(), REPLACE_EXISTING);
                        tryAgain = false;
                    } catch (IOException ex) {
                        if (JOptionPane.showConfirmDialog(null, ex.getMessage() + "\nDo you want to retry?",
                                "Error writing file", JOptionPane.YES_NO_OPTION) != JOptionPane.YES_OPTION) {
                            return false;
                        }
                        tryAgain = true;

                    }
                } while (tryAgain);
                return true;
            }
        } catch (IOException ex) {
            log.warn(ex.getMessage());
            return false;
        }
    }
}
