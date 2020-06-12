package com.ykrkn.xl2json;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.github.jankroken.commandline.CommandLineParser;
import com.github.jankroken.commandline.OptionStyle;
import com.github.jankroken.commandline.domain.InvalidCommandLineException;
import com.github.jankroken.commandline.domain.InvalidOptionConfigurationException;
import com.github.jankroken.commandline.domain.UnrecognizedSwitchException;
import lombok.extern.slf4j.Slf4j;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

@Slf4j
public class Main {

    private static final ObjectMapper OM = new ObjectMapper();

    public static void main(String[] args) {
        Arguments arguments = null;

        try {
            arguments = CommandLineParser.parse(Arguments.class, args, OptionStyle.LONG_OR_COMPACT);
        } catch (InvalidCommandLineException | InvalidOptionConfigurationException | UnrecognizedSwitchException | IllegalAccessException | InstantiationException | InvocationTargetException e) {
            Arguments.printHelp();
            System.exit(1);
        }

        if (arguments.isPrintHelp()) {
            Arguments.printHelp();
            System.exit(0);
        }

        log.info(arguments.toString());

        final String filename = arguments.getFilename();

        final Path path = Paths.get(filename).toAbsolutePath();
        log.info("Parse file " + path);

        MatrixExcelCallback callback = new MatrixExcelCallback(filename, arguments.isParseEmptyRows());
        
        ExcelParser excelParser = new ExcelParser(arguments.toParserOptions());

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
