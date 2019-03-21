/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.save;

import de.bundesbank.jdemetra.exceladapter.Data;
import ec.nbdemetra.sa.MultiProcessingManager;
import ec.nbdemetra.sa.SaBatchUI;
import ec.nbdemetra.ws.actions.AbstractViewAction;
import java.util.Arrays;
import java.util.List;
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

public final class SelectionSaveExcel extends AbstractViewAction<SaBatchUI> {

    public SelectionSaveExcel() {
        super(SaBatchUI.class);
        putValue(NAME, Bundle.CTL_SelectionSaveExcel());
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
        List<Data> items = Arrays.asList(new Data(context.getName(), Arrays.asList(context.getSelection())));
        Thread t = new Thread(new SaveRunnable("Save Selection to Excel", items), "SaveSelectionToExcel-Thread");
        t.start();
    }
}
