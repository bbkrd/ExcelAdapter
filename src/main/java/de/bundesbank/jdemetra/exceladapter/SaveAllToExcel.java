/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import ec.nbdemetra.sa.MultiProcessingDocument;
import ec.nbdemetra.sa.MultiProcessingManager;
import ec.nbdemetra.ws.IWorkspaceItemManager;
import ec.nbdemetra.ws.Workspace;
import ec.nbdemetra.ws.WorkspaceFactory;
import ec.nbdemetra.ws.WorkspaceItem;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import org.netbeans.api.progress.ProgressHandle;
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
        Thread t = new Thread(() -> {
            ProgressHandle progressHandle = ProgressHandle.createHandle("Save all Information to Excel");
            progressHandle.start();

            JFileChooser chooser = ExcelFileChooser.INSTANCE;
            if (chooser.showSaveDialog(null) != JFileChooser.APPROVE_OPTION) {
                progressHandle.finish();
                return;
            }

            File file = chooser.getSelectedFile();
            Workspace workspace = WorkspaceFactory.getInstance().getActiveWorkspace();
            IWorkspaceItemManager mgr = WorkspaceFactory.getInstance().getManager(MultiProcessingManager.ID);
            if (mgr != null) {
                List<WorkspaceItem<MultiProcessingDocument>> list = workspace.searchDocuments(mgr.getItemClass());
                SaveData sd = new SaveData();
                list.stream().forEach((item) -> {
                    sd.addData(item.getDisplayName(), item.getElement().getCurrent());
                });
                if (sd.save(file)) {
                    if (JOptionPane.showConfirmDialog(null, "Do you want to set the saved file as new input for ConCur?", "Input for ConCur", JOptionPane.YES_NO_OPTION)
                            == JOptionPane.YES_OPTION) {
                        String filePath = file.getAbsolutePath();
                        ExcelMetaDataHelper metaDataHelper = new ExcelMetaDataHelper(filePath);
                        list.stream()
                                .forEach(item -> {
                                    String multidocName = item.getDisplayName();
                                    item.getElement().getCurrent().forEach(saItem -> metaDataHelper.overrideMetaData(saItem, multidocName));
                                });
                    }
                }
            }

            progressHandle.finish();
        }, "SaveAllToExcel-Thread");
        t.start();

    }
}
