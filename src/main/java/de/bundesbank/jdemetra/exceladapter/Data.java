/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package de.bundesbank.jdemetra.exceladapter;

import ec.tss.sa.SaItem;
import java.util.List;

/**
 *
 * @author Thomas Witthohn
 */
@lombok.Value
public class Data {

    String name;
    List<SaItem> current;

    public Data(String name, List<SaItem> current) {
        this.name = shortenName(name);
        this.current = current;
    }

    private String shortenName(String name) {
        return name.length() > 20 ? name.substring(0, 15) + "..." : name;
    }
}
