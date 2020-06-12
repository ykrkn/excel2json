package com.ykrkn.xl2json;

import com.github.jankroken.commandline.annotations.*;
import lombok.Getter;
import lombok.ToString;

import java.util.ArrayList;
import java.util.List;

@ToString
public class Arguments {

    @Getter
    private String filename;

    private boolean parseHiddenRows;

    private boolean parseHiddenCols;

    @Getter
    private boolean parseEmptyRows;

    @Getter
    @ToString.Exclude
    private boolean printHelp;

    @Option
    @LongSwitch("filename")
    @ShortSwitch("f")
    @SingleArgument
    @Required
    public void setFilename(String value) {
        this.filename = value;
    }

    @Option
    @LongSwitch("parse-hidden-rows")
    @ShortSwitch("r")
    @Toggle(true)
    public void setParseHiddenRows(boolean value) {
        this.parseHiddenRows = value;
    }

    @Option
    @LongSwitch("parse-empty-rows")
    @ShortSwitch("e")
    @Toggle(true)
    public void setParseEmptyRows(boolean value) {
        this.parseEmptyRows = value;
    }

    @Option
    @LongSwitch("parse-hidden-cols")
    @ShortSwitch("c")
    @Toggle(true)
    public void setParseHiddenCols(boolean value) {
        this.parseHiddenCols = value;
    }

    @Option
    @LongSwitch("help")
    @ShortSwitch("h")
    @Toggle(true)
    public void setPrintHelp(boolean value) {
        this.printHelp = value;
    }

    public SheetParser.Options[] toParserOptions() {
        List<SheetParser.Options> opts = new ArrayList<>();
        if (parseHiddenCols) opts.add(SheetParser.Options.PARSE_HIDDEN_COLS);
        if (parseHiddenRows) opts.add(SheetParser.Options.PARSE_HIDDEN_ROWS);
        return opts.toArray(new SheetParser.Options[0]);
    }

    public static void printHelp() {
        System.out.println("java -jar xl2json.jar --filename|-f <filename.xlsx>");
        System.out.println("\tExcel workbook xls/xlsx will be converted into <filename.xlsx.json>");
        System.out.println("options:");
        System.out.println("\t--parse-hidden-cols|-c toggle parsing hidden columns");
        System.out.println("\t--parse-hidden-rows|-r toggle parsing hidden rows");
        System.out.println("\t--parse-empty-rows|-e toggle parsing empty rows (all the row's cells are empty)");
        System.out.println("\t--help|-h prints this message");
    }
}