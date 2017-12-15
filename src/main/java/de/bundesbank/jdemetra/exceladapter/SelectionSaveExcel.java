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
import java.util.Arrays;
import javax.swing.JFileChooser;
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
@NbBundle.Messages("CTL_SelectionSaveExcel=Save Excel")

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
            if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
                progressHandle.finish();
                return;
            }

            SaItem[] selection = context.getSelection();
            SaveData x = new SaveData();
            x.addData(context.getName(), Arrays.asList(selection));
            x.save(chooser.getSelectedFile());
            progressHandle.finish();
        }, "SaveSelectionToExcel-Thread");
        t.setDaemon(false);
        t.start();
    }
}
