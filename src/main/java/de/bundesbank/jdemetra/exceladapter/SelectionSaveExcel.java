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
        SaItem[] selection = context.getSelection();
        SaveData x = new SaveData();
        x.addData(context.getName(), Arrays.asList(selection));
        x.save();
        context.setSelection(new SaItem[0]);
        context.setSelection(selection);
    }
}
