package com.ykrkn.xl2json;

import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.extern.slf4j.Slf4j;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * java -jar luxmsbi-excel-parser-4.0.jar <filename>
 */
@Slf4j
public class Main {

    private static final ObjectMapper OM = new ObjectMapper();

    public static void main(String[] args) {
        if (args.length < 1) {
            log.info("Input file is a must");
            return;
        }
        final String filename = args[0];

        final Path path = Paths.get(args[0]).toAbsolutePath();
        log.info("Parse file " + path);

        MatrixExcelCallback callback = new MatrixExcelCallback(filename);
        ExcelParser excelParser = new ExcelParser();


        try (final InputStream in = new FileInputStream(path.toFile())) {
            excelParser.parse(in, callback);
            List<SheetResult> result = callback.getSheets();
            Path sink = Paths.get(path.toString() + ".json");
            OM.writeValue(sink.toFile(), result);
            log.info("Saved into " + sink);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
