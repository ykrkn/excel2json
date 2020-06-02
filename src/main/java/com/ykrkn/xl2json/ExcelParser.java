package com.ykrkn.xl2json;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

@Slf4j
public class ExcelParser {
    
    public void parse(InputStream is, ExcelParserCallback callback) {
        try {
            Workbook workbook = WorkbookFactory.create(is);
            parse(workbook, callback);
            callback.onComplete();
        } catch (IOException | InvalidFormatException e) {
            callback.onError(e);
            throw new RuntimeException(e);
        }
    }

    public void parse(Workbook workbook, ExcelParserCallback callback) {
        Iterator<Sheet> iter = workbook.sheetIterator();
        while (iter.hasNext()) {
            Sheet sheet = iter.next();

            if(callback.shouldBreak()) break;
            log.info("Recognize sheet:{}", sheet.getSheetName());

            SheetParser sheetParser = new SheetParser(sheet,
                    SheetParser.Options.DATE_FORMAT_FROM_CELL,
                    SheetParser.Options.SKIP_HIDDEN_COLS,
                    SheetParser.Options.SKIP_HIDDEN_ROWS
            );

            callback.onSheet(sheet.getSheetName());
            sheetParser.parse(callback);
            callback.onSheetComplete();
        }
    }

}
