package com.ykrkn.xl2json;

import java.util.ArrayList;
import java.util.List;

public class MatrixExcelCallback implements ExcelParserCallback {

    private final String workbookName;

    private final List<SheetResult> sheets = new ArrayList<>();

    private String currentSheetName;

    private SheetMatrix currentSheet;

    public MatrixExcelCallback(String workbookName) {
        this.workbookName = workbookName;
    }

    @Override
    public void onRowStart(int rowIndex) {

    }

    @Override
    public void onRowEnd(int rowIndex) {

    }

    @Override
    public void onCell(int rowIndex, String colName, Object value) {
        currentSheet.put(rowIndex, colName, value);
    }

    @Override
    public void onSheet(String sheetName) {
        currentSheetName = sheetName;
        currentSheet = new SheetMatrix();
    }

    @Override
    public void onSheetComplete() {
        sheets.add(new SheetResult(workbookName, currentSheetName, currentSheet));
    }

    public List<SheetResult> getSheets() {
        return sheets;
    }
}
