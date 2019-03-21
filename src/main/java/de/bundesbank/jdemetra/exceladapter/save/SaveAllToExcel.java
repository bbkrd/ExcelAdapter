/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.save;

import de.bundesbank.jdemetra.exceladapter.Data;
import ec.nbdemetra.sa.MultiProcessingDocument;
import ec.nbdemetra.sa.MultiProcessingManager;
import ec.nbdemetra.ws.IWorkspaceItemManager;
import ec.nbdemetra.ws.Workspace;
import ec.nbdemetra.ws.WorkspaceFactory;
import ec.nbdemetra.ws.WorkspaceItem;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.List;
import org.openide.awt.ActionID;
import org.openide.awt.ActionReference;
import org.openide.awt.ActionRegistration;
import org.openide.util.NbBundle.Messages;

@ActionID(
        category = "SA",
        id = "de.bundesbank.jdemetra.exceladapter.SaveAllToExcel"
)
@ActionRegistration(
        displayName = "#CTL_SaveAllToExcel"
)
@ActionReference(path = "Menu/Statistical methods", position = 3333)
@Messages("CTL_SaveAllToExcel=Save all MD to Excel")
public final class SaveAllToExcel implements ActionListener {

    @Override
    public void actionPerformed(ActionEvent e) {
        Workspace workspace = WorkspaceFactory.getInstance().getActiveWorkspace();
        IWorkspaceItemManager mgr = WorkspaceFactory.getInstance().getManager(MultiProcessingManager.ID);
        if (mgr != null) {
            List<WorkspaceItem<MultiProcessingDocument>> list = workspace.searchDocuments(mgr.getItemClass());
            List<Data> items = new ArrayList<>(list.size());
            int counter = 0;
            for (WorkspaceItem<MultiProcessingDocument> item : list) {
                items.add(new Data(++counter + item.getDisplayName(), item.getElement().getCurrent()));
            }
            Thread t = new Thread(new SaveRunnable("Save all Information to Excel", items), "SaveAllToExcel-Thread");
            t.start();
        }
    }
}
