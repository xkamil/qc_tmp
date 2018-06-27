package com.util;

import com.App;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class WorkbookTools {

    public static List<Entry<String, Cell>> findCellsWithValue(Map<String, Workbook> workbooks, String value, MatchType matchType) {
        List<Entry<String, Cell>> cells = new ArrayList<>();

        workbooks.forEach((fileName, workbook) -> {
            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        String cellValue = CellMapper.toString(cell).trim().toLowerCase();

                        switch (matchType) {
                            case EQUALS:
                                if (cellValue.equalsIgnoreCase(value)) {
                                    cells.add(new Entry<>(fileName, cell));
                                }
                                break;
                            case CONTAINS:
                                if (cellValue.contains(value.toLowerCase())) {
                                    cells.add(new Entry<>(fileName, cell));
                                }
                                break;
                            case ENDS_WITH:
                                if (cellValue.endsWith(value.toLowerCase())) {
                                    cells.add(new Entry<>(fileName, cell));
                                }
                                break;
                            case STARTS_WITH:
                                if (cellValue.startsWith(value.toLowerCase())) {
                                    cells.add(new Entry<>(fileName, cell));
                                }
                                break;
                        }
                    }
                }
            }
        });


        return cells;
    }

    public static class Entry<K, V> {
        private K key;
        private V value;

        public Entry(K key, V value) {
            this.key = key;
            this.value = value;
        }

        public K getKey() {
            return key;
        }

        public V getValue() {
            return value;
        }
    }

}
