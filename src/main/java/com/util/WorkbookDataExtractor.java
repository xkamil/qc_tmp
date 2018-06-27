package com.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.HashMap;
import java.util.Map;


public final class WorkbookDataExtractor {

    public static Map<String, String> getRowWithHeaders(Cell cell) {
        Map<String, String> mappedRow = new HashMap<>();

        Row headerRow = cell.getSheet().getRow(0);
        Row dataRow = cell.getRow();


        for (Cell cl : dataRow) {
            int clIdx = cl.getColumnIndex();

            String data = CellMapper.toString(cl);
            String header = CellMapper.toString(headerRow.getCell(clIdx)).trim().toLowerCase();

            mappedRow.put(header, data);
        }

        return mappedRow;

    }

    public static String getCellHeader(Cell cell) {
        int cellIdx = cell.getColumnIndex();
        Cell headerCell = cell.getRow().getSheet().getRow(0).getCell(cellIdx);

        return CellMapper.toString(headerCell);
    }
}
