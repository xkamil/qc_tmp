package com.util;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import static junit.framework.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class WorkbookParserTest {

    private static final Path VALID_DIRECTORY_PATH = Paths.get("src", "test", "resources", "res");
    private static final Path INVALID_DIRECTORY_PATH = Paths.get("src", "test", "resources", "res2");
    private static final Path FILE_PATH = Paths.get("src", "test", "resources", "res", "sheet1.xlsx");
    private static final List<String> fileNames = Collections.unmodifiableList(Arrays.asList(
            "sheet1.xlsx",
            "sheet2.xlsm",
            "sheet3.xls"
    ));

    @Test
    public void testGetAllWorkbooksReturnsListOfWorkbooks() throws IOException, InvalidFormatException {
        Map<String, Workbook> workbooks = WorkbookReader.getAllWorkbooks(VALID_DIRECTORY_PATH.toString());
        assertEquals(3, workbooks.keySet().size());

        fileNames.forEach(name -> assertTrue(workbooks.keySet().contains(name)));
        workbooks.forEach((k, v) -> assertTrue(v != null));
    }

    @Test(expected = IOException.class)
    public void testGetAllWorkbooksThrowsWhenInvalidDirectoryPath() throws IOException {
        WorkbookReader.getAllWorkbooks(INVALID_DIRECTORY_PATH.toString());
    }

    @Test(expected = IOException.class)
    public void testGetAllWorkbooksThrowsWhenNotDirectoryPath() throws IOException {
        WorkbookReader.getAllWorkbooks(FILE_PATH.toString());
    }

}
