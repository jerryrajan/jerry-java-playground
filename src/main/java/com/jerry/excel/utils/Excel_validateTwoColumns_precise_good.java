package com.jerry.excel.utils;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Excel_validateTwoColumns_precise_good {
    public static void main(String[] args) throws IOException {

        List<String> columnValue1 = getColumnValues("src/main/resources/Excel_validateTwoColumns_precise/file1.xls",
                "sheetA", 2, "status A");
        List<String> columnValue2 = getColumnValues("src/main/resources/Excel_validateTwoColumns_precise/file2.xls",
                "sheetB", 1, "status B");

        System.out.println("checking equivalent : " + compareStatusLists(columnValue1, columnValue2));
    }

    public static boolean compareStatusLists(List<String> list1_status, List<String> list2_status) {
        if (list1_status.size() != list2_status.size()) {
            return false;
        }

        for (int i = 0; i < list1_status.size(); i++) {
            String status1 = list1_status.get(i);
            String status2 = list2_status.get(i);

            if (!areEquivalent(status1, status2)) {
                return false;
            }
        }

        return true;
    }

    public static boolean areEquivalent(String status1, String status2) {
        Map<String, String> equivalenceMap = new HashMap<>();
        equivalenceMap.put("Y", "Yes");
        equivalenceMap.put("N", "No");

        String equivalentStatus1 = equivalenceMap.getOrDefault(status1, status1);
        String equivalentStatus2 = equivalenceMap.getOrDefault(status2, status2);

        return equivalentStatus1.equals(equivalentStatus2);
    }

    public static List<String> getColumnValues(String filePath, String sheetName, int titleRowNumber, String columnName)
            throws IOException {
        List<String> columnValues = new ArrayList<>();

        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fileInputStream)) {

            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet == null) {
                throw new IllegalArgumentException("Sheet '" + sheetName + "' not found in the workbook.");
            }

            Row titleRow = sheet.getRow(titleRowNumber - 1); // Convert 1-based to 0-based row index
            if (titleRow == null) {
                throw new IllegalArgumentException("Title row " + titleRowNumber + " not found.");
            }

            int columnIndex = -1;

            for (Cell cell : titleRow) {
                if (cell.getStringCellValue().trim().equals(columnName)) {
                    columnIndex = cell.getColumnIndex();
                    break;
                }
            }

            if (columnIndex == -1) {
                throw new IllegalArgumentException("Column '" + columnName + "' not found in the title row.");
            }

            for (int i = titleRowNumber; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell cell = row.getCell(columnIndex);
                String value = getCellValueAsString(cell).trim();
                columnValues.add(value);
            }
        }

        return columnValues;
    }

    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
