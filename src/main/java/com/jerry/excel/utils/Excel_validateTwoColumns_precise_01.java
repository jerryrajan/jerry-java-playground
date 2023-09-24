package com.jerry.excel.utils;



import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

public class Excel_validateTwoColumns_precise_01 {
    public static void main(String[] args) throws IOException {

        List<String> columnValue1 = getColumnValues("src/main/resources/Excel_validateTwoColumns_precise/file1.xls",
                "sheetA",
                2,
                "status A");
        List<String> columnValue2 = getColumnValues("src/main/resources/Excel_validateTwoColumns_precise/file2.xls",
                "sheetB",
                1,
                "status B");
        System.out.println("Number of values: " + columnValue1.size());
        System.out.println("list of columns: " + columnValue1);
        System.out.println("Number of values: " + columnValue2.size());
        System.out.println("list of columns: " + columnValue2);

        System.out.println("checking equivalent : " +areEquivalent(columnValue1,columnValue2,
                "appleToApple"));
        System.out.println("checking equivalent : " +areEquivalent(columnValue1,columnValue2,
                "yesToY"));

    }

    public static boolean areEquivalent(List<String> list1, List<String> list2, String validationType) {
        if (list1.size() != list2.size()) {
            System.out.println("Lists have different sizes.");
            return false;
        }

        if ("appleToApple".equalsIgnoreCase(validationType)) {
            boolean result = IntStream.range(0, list1.size())
                    .allMatch(i -> {
                        String value1 = list1.get(i);
                        String value2 = list2.get(i);
                        if (!value1.equals(value2)) {
                            System.out.println("Lists are not equal at index " + i + " (Apple to Apple comparison):");
                            System.out.println("List1: " + value1);
                            System.out.println("List2: " + value2);
                            return false;
                        }
                        return true;
                    });

            if (!result) {
                System.out.println("Apple-to-Apple Comparison: Lists are not equal.");
            } else {
                System.out.println("Apple-to-Apple Comparison: Lists are equal.");
            }

            return result;
        } else if ("yesToY".equalsIgnoreCase(validationType)) {
            Map<String, String> equivalenceMap = new HashMap<>();
            equivalenceMap.put("Y", "Yes");
            equivalenceMap.put("N", "No");

            boolean result = IntStream.range(0, list1.size())
                    .allMatch(i -> {
                        String value1 = list1.get(i);
                        String value2 = list2.get(i);
                        if (!value1.equals(value2) && !equivalenceMap.getOrDefault(value1, value1).equals(value2)) {
                            System.out.println("Lists are not equal at index " + i + " (Yes to Y comparison):");
                            System.out.println("List1: " + value1);
                            System.out.println("List2: " + value2);
                            return false;
                        }
                        return true;
                    });

            if (!result) {
                System.out.println("Yes to Y Comparison: Lists are not equal.");
            } else {
                System.out.println("Yes to Y Comparison: Lists are equal.");
            }

            return result;
        }

        System.out.println("Invalid validation type.");
        return false;
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
