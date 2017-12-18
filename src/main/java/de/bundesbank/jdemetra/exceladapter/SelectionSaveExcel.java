/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import ec.nbdemetra.sa.MultiProcessingManager;
import ec.nbdemetra.sa.SaBatchUI;
import ec.nbdemetra.ws.actions.AbstractViewAction;
import ec.tss.sa.SaItem;
import java.io.File;
import java.util.Arrays;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.netbeans.api.progress.ProgressHandle;
import org.openide.awt.ActionID;
import org.openide.awt.ActionReference;
import org.openide.awt.ActionRegistration;
import org.openide.util.NbBundle;

/**
 *
 * @author Thomas Witthohn
 */
@ActionID(
        category = "SA",
        id = "de.bundesbank.jdemetra.exceladapter.SaveExcel"
)

@ActionRegistration(
        displayName = "#CTL_SelectionSaveExcel",
        lazy = true
)

@ActionReference(path = MultiProcessingManager.LOCALPATH)
@NbBundle.Messages("CTL_SelectionSaveExcel=Save to Excel")

public class SelectionSaveExcel extends AbstractViewAction<SaBatchUI> {

    public SelectionSaveExcel() {
        super(SaBatchUI.class);
        putValue(NAME, Bundle.CTL_SelectionAddExcelMeta());
        refreshAction();
    }

    @Override
    protected void refreshAction() {
        if (context() != null) {
            setEnabled(context().getSelectionCount() > 0);
        }
    }

    @Override
    protected void process(SaBatchUI context) {
        Thread t = new Thread(() -> {
            ProgressHandle progressHandle = ProgressHandle.createHandle("Save Selection to Excel");
            progressHandle.start();

            JFileChooser chooser = ExcelFileChooser.INSTANCE;
            if (chooser.showSaveDialog(null) != JFileChooser.APPROVE_OPTION) {
                progressHandle.finish();
                return;
            }
            File file = chooser.getSelectedFile();
            SaItem[] selection = context.getSelection();
            String multiDocName = context.getName();
            SaveData sd = new SaveData();
            sd.addData(multiDocName, Arrays.asList(selection));

            if (sd.save(file)) {
                if (JOptionPane.showConfirmDialog(null, "Do you want to set the saved file as new input for ConCur?", "Input for ConCur", JOptionPane.YES_NO_OPTION)
                        == JOptionPane.YES_OPTION) {
                    String filePath = file.getAbsolutePath();
                    ExcelMetaDataHelper metaDataHelper = new ExcelMetaDataHelper(filePath);
                    for (SaItem saItem : selection) {
                        metaDataHelper.overrideMetaData(saItem, multiDocName);
                    }
                }
            }
            progressHandle.finish();
        }, "SaveSelectionToExcel-Thread");
        t.setDaemon(false);
        t.start();
    }
}
