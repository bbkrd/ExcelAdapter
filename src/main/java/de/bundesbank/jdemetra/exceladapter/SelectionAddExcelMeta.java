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
        id = "de.bundesbank.jdemetra.exceladapter.AddExcelMeta"
)

@ActionRegistration(
        displayName = "#CTL_SelectionAddExcelMeta",
        lazy = true
)

@ActionReference(path = MultiProcessingManager.LOCALPATH)
@NbBundle.Messages("CTL_SelectionAddExcelMeta=Add Excel MetaData")

public class SelectionAddExcelMeta extends AbstractViewAction<SaBatchUI> {

    public SelectionAddExcelMeta() {
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
        ExcelMetaDataHelper metaDataHelper = new ExcelMetaDataHelper();
        SaItem[] selection = context.getSelection();
        String multiDocName = context.getName();

        for (SaItem saItem : selection) {
            metaDataHelper.addMetaData(saItem, multiDocName);
        }
        context.setSelection(new SaItem[0]);
        context.setSelection(selection);
    }

}
