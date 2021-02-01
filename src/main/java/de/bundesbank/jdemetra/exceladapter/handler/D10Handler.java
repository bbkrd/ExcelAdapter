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
public class D10Handler extends AbstractHandler {

    public static final String SUFFIX = "-D10";
    private static final String NAME = "D10",
            OPTION_ID = "exceladapter.d10";
    private static final boolean DEFAULT = true;

    @Override
    protected List<Ts> extractData(List<SaItem> items) {
        List<Ts> list = new ArrayList<>();

        items.forEach((saItem) -> {
            String name = ExcelMetaDataHelper.getItemName(saItem);;
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
