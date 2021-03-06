/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 *
 * @author Deutsche Bundesbank
 */
public final class ExcelFileChooser extends JFileChooser {

    public static final ExcelFileChooser INSTANCE = new ExcelFileChooser();

    private ExcelFileChooser() {
        super();
        setDialogTitle("Save as");
        setAcceptAllFileFilterUsed(false);
        FileFilter filter = new FileNameExtensionFilter("XLSX files", "xlsx");
        addChoosableFileFilter(filter);
    }

    @Override
    public void approveSelection() {
        File f = getSelectedFile();
        if (!f.exists() || getDialogType() != SAVE_DIALOG) {
            super.approveSelection();
            return;
        }
        int result = JOptionPane.showConfirmDialog(this, "The file already exists! Do you want to overwrite it?",
                "Existing file", JOptionPane.YES_NO_CANCEL_OPTION);
        switch (result) {
            case JOptionPane.YES_OPTION:
                super.approveSelection();
                return;
            case JOptionPane.NO_OPTION:
            case JOptionPane.CLOSED_OPTION:
                return;
            case JOptionPane.CANCEL_OPTION:
                cancelSelection();
        }
    }

    @Override
    public void setSelectedFile(File file) {
        if (file != null) {
            String withExtension = file.getAbsolutePath();
            if (!withExtension.toLowerCase().endsWith(".xlsx")) {
                withExtension += ".xlsx";
            }
            super.setSelectedFile(new File(withExtension));
            return;
        }
        super.setSelectedFile(file);
    }

}
