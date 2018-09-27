/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.options;

import java.beans.PropertyChangeListener;
import java.beans.PropertyChangeSupport;
import javax.swing.JComponent;
import javax.swing.SwingUtilities;
import org.netbeans.spi.options.OptionsPanelController;
import org.openide.util.HelpCtx;
import org.openide.util.Lookup;

@OptionsPanelController.SubRegistration(
        location = "Demetra",
        displayName = "#AdvancedOption_DisplayName_ExcelAdapter",
        keywords = "#AdvancedOption_Keywords_ExcelAdapter",
        keywordsCategory = "Advanced/ExcelAdapter"
)
@org.openide.util.NbBundle.Messages({"AdvancedOption_DisplayName_ExcelAdapter=ExcelAdapter", "AdvancedOption_Keywords_ExcelAdapter=ConCur,Excel"})
public final class ExcelAdapterOptionsPanelController extends OptionsPanelController {

    private ExcelAdapterPanel panel;
    private final PropertyChangeSupport pcs = new PropertyChangeSupport(this);
    private boolean changed;

    public static final String A6 = "exceladapter.a6";
    public static final String D10 = "exceladapter.d10";
    public static final String D11 = "exceladapter.d11";
    public static final String D12 = "exceladapter.d12";
    public static final String D13 = "exceladapter.d13";

    @Override
    public void update() {
        getPanel().load();
        changed = false;
    }

    @Override
    public void applyChanges() {
        SwingUtilities.invokeLater(() -> {
            getPanel().store();
            changed = false;
        });
    }

    @Override
    public void cancel() {
        // need not do anything special, if no changes have been persisted yet
    }

    @Override
    public boolean isValid() {
        return getPanel().valid();
    }

    @Override
    public boolean isChanged() {
        return changed;
    }

    @Override
    public HelpCtx getHelpCtx() {
        return null; // new HelpCtx("...ID") if you have a help set
    }

    @Override
    public JComponent getComponent(Lookup masterLookup) {
        return getPanel();
    }

    @Override
    public void addPropertyChangeListener(PropertyChangeListener l) {
        pcs.addPropertyChangeListener(l);
    }

    @Override
    public void removePropertyChangeListener(PropertyChangeListener l) {
        pcs.removePropertyChangeListener(l);
    }

    private ExcelAdapterPanel getPanel() {
        if (panel == null) {
            panel = new ExcelAdapterPanel(this);
        }
        return panel;
    }

    void changed() {
        if (!changed) {
            changed = true;
            pcs.firePropertyChange(OptionsPanelController.PROP_CHANGED, false, true);
        }
        pcs.firePropertyChange(OptionsPanelController.PROP_VALID, null, null);
    }

}
