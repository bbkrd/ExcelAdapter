/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

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
import org.openide.util.Pair;
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = IHandler.class)
public class A1AHandler implements IHandler {

    @Override
    public Pair<String, List<Ts>> extractData(String multidocName, List<SaItem> items) {
        List<Ts> list = new ArrayList<>();

        items.forEach((saItem) -> {
            String name = saItem.getRawName().isEmpty() ? saItem.getTs().getRawName() : saItem.getRawName();
            CompositeResults result = saItem.process();
            if (result == null || result.getData(ModellingDictionary.MODE, DecompositionMode.class) == null) {
                return;
            }
            TsData tsData = result.getData(X11Kernel.A1a, TsData.class);
            if (tsData != null) {
                Ts ts = TsFactory.instance.createTs(name, null, tsData);
                list.add(ts);
            }
        });

        return Pair.of(multidocName + "-A1A", list);

    }

    @Override
    public boolean isEnabled() {
        return true;
    }

}
