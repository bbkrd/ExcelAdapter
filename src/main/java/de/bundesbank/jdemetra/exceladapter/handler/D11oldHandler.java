/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import de.bbk.concur.util.SeasonallyAdjusted_Saved;
import de.bundesbank.jdemetra.exceladapter.ExcelMetaDataHelper;
import ec.tss.Ts;
import ec.tss.sa.SaItem;
import ec.tstoolkit.algorithm.CompositeResults;
import java.util.ArrayList;
import java.util.List;
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = AbstractHandler.class)
public class D11oldHandler extends AbstractHandler {

    private static final String SUFFIX = "-D11(old)";
    private static final String NAME = "D11 old",
            OPTION_ID = "exceladapter.d11old";
    private static final boolean DEFAULT = false;
    @Override
    protected List<Ts> extractData(List<SaItem> items) {
        List<Ts> list = new ArrayList<>();

        items.forEach((saItem) -> {
            String name = ExcelMetaDataHelper.getItemName(saItem);
            CompositeResults result = saItem.process();
            if (result == null) {
                return;
            }

            Ts ts = SeasonallyAdjusted_Saved.calcSeasonallyAdjusted(saItem.toDocument());
            list.add(ts.rename(name));
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
