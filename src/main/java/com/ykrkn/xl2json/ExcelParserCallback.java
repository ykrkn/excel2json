package com.ykrkn.xl2json;

public interface ExcelParserCallback {

    void onRowStart(int rowIndex);

    void onRowEnd(int rowIndex);

    void onCell(int rowIndex, String colName, Object value);

    void onSheet(String sheetName);

    void onSheetComplete();

    default void onComplete(){}

    default void onError(Throwable ex){}

    default boolean shouldBreak() {
        return false;
    }
}
