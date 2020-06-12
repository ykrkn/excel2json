package com.ykrkn.xl2json;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.util.Date;

@Slf4j
public class SheetParser {
    
    public enum Options {
        PARSE_HIDDEN_ROWS,
        PARSE_HIDDEN_COLS,
        DATE_FORMAT_FROM_CELL
    }

    private boolean dateFormatFromCell;

    private boolean parseHiddenRows;

    private boolean parseHiddenCols;

    private final Sheet sheet;

    public SheetParser(Sheet sheet, Options ... opts) {
        this.sheet = sheet;
        for (int i = 0; i < opts.length; i++) {
            Options opt = opts[i];
            if (opt == Options.PARSE_HIDDEN_COLS) parseHiddenCols = true;
            if (opt == Options.PARSE_HIDDEN_ROWS) parseHiddenRows = true;
            if (opt == Options.DATE_FORMAT_FROM_CELL) dateFormatFromCell = true;
        }
    }

    public void parse(ExcelParserCallback callback) {
        for(Row row : sheet) {
            if(!parseHiddenRows && row.getZeroHeight()) {
                log.warn("Row [" + row.getRowNum() + "] is hidden - will be skipped");
                continue;
            }

            if(callback.shouldBreak()) break;

            callback.onRowStart(row.getRowNum());

            for (Cell cell : row) {
                if(!parseHiddenCols && sheet.isColumnHidden(cell.getColumnIndex())) {
                    log.warn("Column [" + cell.getColumnIndex() + "] is hidden - will be skipped");
                    continue;
                }

                callback.onCell(
                    cell.getRowIndex(),
                    CellReference.convertNumToColString(cell.getColumnIndex()).toLowerCase(),
                    getCellValue(cell)
                );
            }

            callback.onRowEnd(row.getRowNum());
        }
    }

    private Object getCellValue(Cell cell) {
        if(cell == null) return null;
        try {
            return getCellValue(cell, cell.getCellTypeEnum());
        } catch (RuntimeException e) {
            log.warn(e.getMessage() + ' ' + cellToString(cell, "UNKNOWN"));
            return null;
        }
    }

    private final Object getCellValue(Cell cell, CellType cellType) {
        Object value = null;
        switch(cellType) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return evaluateDate(cell);
                }
                return cell.getNumericCellValue();
            case FORMULA:
                value = evaluateFormula(cell);
                break;
            default:
                value = cell.getStringCellValue().trim();
                if ("".equals(value)) value = null;
        }
        return value;
    }

    // https://stackoverflow.com/questions/34900605/excel-cell-style-issue/34902174#34902174
    private final Object evaluateDate(Cell cell) {
        Date date = cell.getDateCellValue();
        if (!dateFormatFromCell) return date;

        String dateFmt = (cell.getCellStyle().getDataFormat() == 14
                ? "dd/mm/yyyy"
                : cell.getCellStyle().getDataFormatString()
        );
        String result = new CellDateFormatter(dateFmt).format(date);
        return result.replace(";@", ""); // bug in the POI ?
    }

    private final Object evaluateFormula(Cell cell) {
        try {
            CellValue cellValue = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator().evaluate(cell);
            switch (cellValue.getCellTypeEnum()) {
                case NUMERIC: return cellValue.getNumberValue();
                default: return cellValue.getStringValue();
            }
        } catch (RuntimeException e) {
            log.warn(e.getMessage());
            // To prevent the stack overflow
            if (cell.getCachedFormulaResultTypeEnum() == CellType.FORMULA) {
                return cell.getCellFormula();
            } else {
                return getCellValue(cell, cell.getCachedFormulaResultTypeEnum());
            }
        }
    }

    private static String cellToString(Cell cell, String value) {
        return cell.getSheet().getSheetName() + ':' + cell.getAddress().formatAsString() + '=' + value.replace('\n', '\u2165');
    }
}
