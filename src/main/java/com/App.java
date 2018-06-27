package com;

import com.util.MatchType;
import com.util.WorkbookDataExtractor;
import com.util.WorkbookReader;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Paths;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import static com.util.WorkbookTools.findCellsWithValue;


public class App {

    public static void main(String[] args) throws IOException {

        Logger.getLogger("").setLevel(Level.INFO);

        String value = "4";

        Map<String, Workbook> workbooks = WorkbookReader.getAllWorkbooks(Paths.get("sheets").toString());

        findCellsWithValue(workbooks, value, MatchType.STARTS_WITH)
                .forEach(c -> {
                    Map<String,String> row = WorkbookDataExtractor.getRowWithHeaders(c.getValue());

                    System.out.println(String.format("%-15s - UQID: %-7s QC: %s", c.getKey(), row.get("uqid"), row.get("code")));

                });
    }


}
