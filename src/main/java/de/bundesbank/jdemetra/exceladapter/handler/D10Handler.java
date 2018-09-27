/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import de.bundesbank.jdemetra.exceladapter.options.ExcelAdapterOptionsPanelController;
import ec.satoolkit.DecompositionMode;
import ec.satoolkit.x11.X11Kernel;
import ec.tss.Ts;
import ec.tss.TsFactory;
import ec.tss.sa.SaItem;
import ec.tstoolkit.algorithm.CompositeResults;
import ec.tstoolkit.modelling.ModellingDictionary;
import ec.tstoolkit.timeseries.simplets.TsData;
import java.util.ArrayList;
import java.util.List;
import org.openide.util.NbPreferences;
import org.openide.util.Pair;
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = IHandler.class)
public class D10Handler implements IHandler {

    @Override
    public Pair<String, List<Ts>> extractData(String multidocName, List<SaItem> items) {
        List<Ts> list = new ArrayList<>();

        items.forEach((saItem) -> {
            String name = saItem.getRawName().isEmpty() ? saItem.getTs().getRawName() : saItem.getRawName();
            CompositeResults result = saItem.process();
            if (result == null || result.getData(ModellingDictionary.MODE, DecompositionMode.class) == null) {
                return;
            }
            boolean isMultiplicative = result.getData(ModellingDictionary.MODE, DecompositionMode.class).isMultiplicative();
            TsData d10 = result.getData(X11Kernel.D10, TsData.class);
            if (d10 != null) {
                TsData d10a = result.getData(X11Kernel.D10a, TsData.class);
                Ts seasonalfactor = TsFactory.instance.createTs(name, null, isMultiplicative ? d10.update(d10a).times(100) : d10.update(d10a));
                list.add(seasonalfactor);
            }
        });

        return Pair.of(multidocName + "-D10", list);

    }

    @Override
    public boolean isEnabled() {
        return NbPreferences.forModule(ExcelAdapterOptionsPanelController.class).getBoolean(ExcelAdapterOptionsPanelController.D10, true);
    }
}
