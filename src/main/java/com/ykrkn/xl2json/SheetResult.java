package com.ykrkn.xl2json;

import lombok.Getter;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Getter
public class SheetResult {

    private final String w;

    private String s;

    private List<Map> r = new ArrayList<>();

    public SheetResult(String workbookName, String sheetName, SheetMatrix mtx) {
        this.w = workbookName;
        this.s = sheetName;
        for (Integer ri : mtx.getRows()) {
            Map row = new LinkedHashMap<>();
            row.put("_ri", ri);
            row.put("_rn", ri+1);
            for (String ci : mtx.getCols(ri)) {
                Object v = mtx.getCell(ci, ri);
                row.put(ci, v);
            }
            r.add(row);
        }
    }
}
