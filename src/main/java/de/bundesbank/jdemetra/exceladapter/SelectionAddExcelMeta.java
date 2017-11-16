/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import de.bbk.concur.util.SavedTables;
import static de.bundesbank.jdemetra.exceladapter.ExcelConnection.*;
import ec.nbdemetra.sa.MultiProcessingManager;
import ec.nbdemetra.sa.SaBatchUI;
import ec.nbdemetra.ws.actions.AbstractViewAction;
import ec.tss.sa.SaItem;
import ec.tstoolkit.MetaData;
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
@NbBundle.Messages("CTL_SelectionAddExcelMeta=Add Excel Meta")

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
        SaItem[] selection = context.getSelection();
        addMetaData(selection);
        context.setSelection(new SaItem[0]);
        context.setSelection(selection);
    }

    private void addMetaData(SaItem[] items) {
        for (SaItem item : items) {
            MetaData metaData = item.getMetaData();
            if (metaData == null) {
                metaData = new MetaData();
            }
            for (SavedTables.TABLES table : SavedTables.TABLES.values()) {
                metaData.putIfAbsent(METADATA_EXCEL_FILE + table, "");
                metaData.putIfAbsent(METADATA_EXCEL_SHEET + table, "");
                metaData.putIfAbsent(METADATA_EXCEL_SERIES + table, "");

            }
            item.setMetaData(metaData);
        }
    }

}
