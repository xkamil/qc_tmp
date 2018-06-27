package com.util;

import org.apache.poi.ss.usermodel.Cell;

public final class CellMapper {
    public static String toString(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case STRING:
                return cell.getStringCellValue();
            case ERROR:
                return String.valueOf(cell.getErrorCellValue());
            default:
                return "";
        }
    }
}
