/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import de.bbk.concur.util.SavedTables;
import de.bbk.concur.util.TsData_Saved;
import de.bundesbank.jdemetra.exceladapter.ExcelMetaDataHelper;
import ec.tss.Ts;
import ec.tss.sa.SaItem;
import java.util.List;
import java.util.stream.Collectors;
import org.openide.util.lookup.ServiceProvider;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceProvider(service = AbstractHandler.class)
public class D10oldHandler extends AbstractHandler {

    private static final String SUFFIX = "-D10(old)";
    private static final String NAME = "D10 old",
            OPTION_ID = "exceladapter.d10old";
    private static final boolean DEFAULT = false;

    @Override
    protected List<Ts> extractData(List<SaItem> items) {
        return items.stream()
                .map(item -> TsData_Saved.convertMetaDataToTs(item.getMetaData(), SavedTables.SEASONALFACTOR).rename(ExcelMetaDataHelper.getItemName(item)))
                .collect(Collectors.toList());
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
