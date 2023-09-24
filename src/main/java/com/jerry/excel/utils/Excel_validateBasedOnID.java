package com.jerry.excel.utils;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Excel_validateBasedOnID {
    public static void main(String[] args) throws IOException {


        List<String> list1_id = getColumnValues("src/main/resources/Excel_validateTwoColumns_basedOnID/file1.xls",
                "sheetA", 2, "id");
        List<String> list1_status = getColumnValues("src/main/resources/Excel_validateTwoColumns_basedOnID/file1.xls",
                "sheetA", 2, "status A");
        List<String> list2_id = getColumnValues("src/main/resources/Excel_validateTwoColumns_basedOnID/file2.xls",
                "sheetB", 1, "id");
        List<String> list2_status = getColumnValues("src/main/resources/Excel_validateTwoColumns_basedOnID/file2.xls",
                "sheetB", 1, "status B");

        System.out.println("list of columns: " + list1_id);
        System.out.println("list of columns: " + list1_status);
        System.out.println("list of columns: " + list2_id);
        System.out.println("list of columns: " + list2_status);

        System.out.println("-----------------------------");
        boolean areStatusesEqual = compareIdStatuses(list1_id,list1_status,list2_id,list2_status);

        if (areStatusesEqual) {
            System.out.println("Statuses are the same for the common IDs.");
        } else {
            System.out.println("Statuses are not the same for the common IDs.");
        }



    }


    public static boolean compareIdStatuses(List<String> list1_id, List<String> list1_status, List<String> list2_id, List<String> list2_status) {
        boolean statusesEqual = true;

        for (int i = 0; i < list1_id.size(); i++) {
            String id = list1_id.get(i);
            String status1 = list1_status.get(i);

            int indexInList2 = list2_id.indexOf(id);

            if (indexInList2 == -1) {
                // ID not found in list2
                statusesEqual = false;
                System.out.println("ID '" + id + "' not found in list2.");
            } else {
                String status2 = list2_status.get(indexInList2);

                if (!areEquivalent(status1, status2)) {
                    statusesEqual = false;
                    System.out.println("ID '" + id + "' has different statuses in list1 ('" + status1 + "') and list2 ('" + status2 + "').");
                }
            }
        }

        return statusesEqual;
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
