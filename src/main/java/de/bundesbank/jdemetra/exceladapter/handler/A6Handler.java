/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import de.bundesbank.jdemetra.exceladapter.ExcelMetaDataHelper;
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
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = AbstractHandler.class)
public class A6Handler extends AbstractHandler {
    public static final String SUFFIX = "-A6A7";
    private static final String NAME = "A6/A7",
            OPTION_ID = "exceladapter.a6";
    private static final boolean DEFAULT = true;

    @Override
    protected List<Ts> extractData(List<SaItem> items) {
        List<Ts> list = new ArrayList<>();

        items.forEach((saItem) -> {
            String name = ExcelMetaDataHelper.getItemName(saItem);
            CompositeResults result = saItem.process();
            if (result == null || result.getData(ModellingDictionary.MODE, DecompositionMode.class) == null) {
                return;
            }
            boolean isMultiplicative = result.getData(ModellingDictionary.MODE, DecompositionMode.class).isMultiplicative();
            TsData calFactorData;
            TsData a6 = result.getData(X11Kernel.A6, TsData.class);
            TsData a7 = result.getData(X11Kernel.A7, TsData.class);
            if (a6 == null) {
                calFactorData = a7;
            } else {
                if (isMultiplicative) {
                    calFactorData = a6.times(a7);
                } else {
                    calFactorData = a6.plus(a7);
                }
            }
            if (calFactorData != null) {
                Ts calFactor = TsFactory.instance.createTs(name, null, isMultiplicative ? calFactorData.times(100) : calFactorData);
                list.add(calFactor);
            }
        });

        return list;

    }

    @Override
    public String getSuffix() {
        return SUFFIX;
    }

    @Override
    public String getName() {
        return NAME;
    }

    @Override
    public String getOptionID() {
        return OPTION_ID;
    }

    @Override
    public boolean getDefault() {
        return DEFAULT;
    }

}
