/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.save;

import de.bundesbank.jdemetra.exceladapter.Data;
import de.bundesbank.jdemetra.exceladapter.ExcelFileChooser;
import de.bundesbank.jdemetra.exceladapter.ExcelMetaDataHelper;
import java.io.File;
import java.util.List;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.netbeans.api.progress.ProgressHandle;

/**
 *
 * @author Thomas Witthohn
 */
public class SaveRunnable implements Runnable {

    private final String name;
    private final List<Data> items;

    public SaveRunnable(String name, List<Data> items) {
        this.name = name;
        this.items = items;
    }

    @Override
    public void run() {
        ProgressHandle progressHandle = ProgressHandle.createHandle(name);
        progressHandle.start();
        try {
            JFileChooser chooser = ExcelFileChooser.INSTANCE;
            if (chooser.showSaveDialog(null) != JFileChooser.APPROVE_OPTION) {
                return;
            }
            File file = chooser.getSelectedFile();
            Save sd = new Save();
            sd.addDatas(items);

            if (sd.save(file)) {
                if (JOptionPane.showConfirmDialog(null, "Do you want to set the saved file as new input for ConCur?", "Input for ConCur", JOptionPane.YES_NO_OPTION)
                        == JOptionPane.YES_OPTION) {
                    String filePath = file.getAbsolutePath();
                    ExcelMetaDataHelper metaDataHelper = new ExcelMetaDataHelper(filePath);
                    metaDataHelper.overrideMetaDatas(items);
                }
            }
        } finally {
            progressHandle.finish();
        }

    }
}
