package com.ykrkn.xl2json;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.Iterator;

@Slf4j
public class ExcelParser {

    private SheetParser.Options[] options = new SheetParser.Options[] {
        SheetParser.Options.DATE_FORMAT_FROM_CELL
    };

    public ExcelParser(SheetParser.Options ... options) {
        SheetParser.Options[] ol = Arrays.copyOf(this.options, this.options.length+options.length);
        for (int i = 0; i < options.length; i++) {
            ol[this.options.length+i] = options[i];
        }
        this.options = ol;
    }

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

            SheetParser sheetParser = new SheetParser(sheet, options);

            callback.onSheet(sheet.getSheetName());
            sheetParser.parse(callback);
            callback.onSheetComplete();
        }
    }

}
