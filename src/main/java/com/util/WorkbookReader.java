package com.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.logging.Logger;
import java.util.stream.Collectors;

import static java.util.Arrays.asList;

public final class WorkbookReader {

    private static final Logger LOGGER = Logger.getLogger(WorkbookReader.class.getName());
    private static final List<String> MAP_TO_XSSF = Collections.unmodifiableList(asList(".xlsx", ".xlsm"));
    private static final List<String> MAP_TO_HSSF = Collections.unmodifiableList(asList(".xls"));


    public static Map<String, Workbook> getAllWorkbooks(String location) throws IOException {
        Path dirPath = Paths.get(location);
        Map<String, Workbook> workbooks = new HashMap<>();
        List<File> files = getAllFilesFrom(dirPath);

        files.forEach(f -> fileToWorkbook(f).ifPresent(w -> workbooks.put(f.getName(), w)));

        return workbooks;
    }


    private static List<File> getAllFilesFrom(Path dirPath) throws IOException {
        if (!Files.isDirectory(dirPath)) {
            throw new IOException(String.format("%s is not a directory", dirPath));
        }

        return Files.list(dirPath)
                .map(Path::toFile)
                .collect(Collectors.toList());

    }

    private static Optional<Workbook> fileToWorkbook(File file) {
        Workbook workbook = null;

        try {
            if (MAP_TO_HSSF.contains(getFileExtension(file.getName()))) {
                workbook = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(file)));
            } else if (MAP_TO_XSSF.contains(getFileExtension(file.getName()))) {
                workbook = new XSSFWorkbook(file);
            }
        } catch (Exception ex) {
            LOGGER.warning(ex.getMessage() + ", File: " + file.getName());
        }

        return Optional.ofNullable(workbook);
    }

    private static String getFileExtension(String fileName) {
        return fileName.substring(fileName.lastIndexOf("."));
    }
}
