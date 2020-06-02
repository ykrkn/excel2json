package com.ykrkn.xl2json;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Collectors;

public class SheetMatrix {

    private final Map<Integer, Map<String, Object>> cells = new LinkedHashMap<>();

    public Object getCell(String ci, Integer ri) {
        return cells.get(ri).get(ci);
    }

    public Iterable<Integer> getRows() {
        return cells.entrySet().stream()
                .map(Map.Entry::getKey)
                .collect(Collectors.toList());
    }

    public Iterable<String> getCols(Integer rowIndex) {
        return cells.get(rowIndex).keySet();
    }

    public void put(int rowIndex, String colName, Object value) {
        Map<String, Object> row = cells.computeIfAbsent(rowIndex, key -> new LinkedHashMap<>());
        row.put(colName, value);
    }
}
