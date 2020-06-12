package com.ykrkn.xl2json;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

import java.util.*;

@Getter
@Slf4j
public class SheetResult {

    private final String w;

    private String s;

    private List<Map> r = new ArrayList<>();

    public SheetResult(String workbookName, String sheetName, SheetMatrix mtx, boolean parseEmptyRows) {
        this.w = workbookName;
        this.s = sheetName;
        for (Integer ri : mtx.getRows()) {
            Map row = new LinkedHashMap<>();

            for (String ci : mtx.getCols(ri)) {
                Object v = mtx.getCell(ci, ri);
                row.put(ci, v);
            }

            if (!parseEmptyRows) {
                if (0 == row.values().stream().filter(Objects::nonNull).count()) {
                    log.debug("Sheet [{}] - empty row [{}] will be skipped", sheetName, (ri+1));
                    continue;
                }
            }

            row.put("_ri", ri);
            row.put("_rn", ri+1);

            r.add(row);
        }
    }
}
