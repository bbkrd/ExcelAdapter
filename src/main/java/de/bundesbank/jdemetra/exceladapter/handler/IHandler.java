/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter.handler;

import ec.tss.Ts;
import ec.tss.sa.SaItem;
import ec.tstoolkit.design.ServiceDefinition;
import java.util.List;
import org.openide.util.Pair;

/**
 *
 * @author Thomas Witthohn
 */
@ServiceDefinition
public interface IHandler {

    Pair<String, List<Ts>> extractData(String multidocName, List<SaItem> items);

    boolean isEnabled();

}
