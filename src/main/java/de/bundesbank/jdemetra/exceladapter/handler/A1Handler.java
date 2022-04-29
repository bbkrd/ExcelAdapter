/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import de.bundesbank.jdemetra.exceladapter.ExcelMetaDataHelper;
import ec.satoolkit.x11.X11Kernel;
import ec.tss.Ts;
import ec.tss.TsFactory;
import ec.tss.sa.SaItem;
import ec.tstoolkit.algorithm.CompositeResults;
import ec.tstoolkit.timeseries.simplets.TsData;
import java.util.ArrayList;
import java.util.List;
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = AbstractHandler.class)
public class A1Handler extends AbstractHandler {

    public static final String SUFFIX = "-A1";
    private static final String NAME = "A1",
            OPTION_ID = "exceladapter.a1";
    private static final boolean DEFAULT = true;

    @Override
    protected List<Ts> extractData(List<SaItem> items) {
        List<Ts> list = new ArrayList<>();

        items.forEach((saItem) -> {
            String name = ExcelMetaDataHelper.getItemName(saItem);
            CompositeResults result = saItem.process();
            if (result == null) {
                return;
            }
            TsData tsData = result.getData(X11Kernel.A1, TsData.class);
            if (tsData != null) {
                Ts ts = TsFactory.instance.createTs(name, null, tsData);
                list.add(ts);
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
