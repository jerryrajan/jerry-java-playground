package com.jerry.excel.utils;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ValidateTwoExcelColumns {

    public static void main(String[] args) {

        try {
            boolean validationSuccessful = validateColumns("src/main/resources/excelFiles/file1.xls", "status A",
                    "src/main/resources/excelFiles/file2.xls", "status B","specialValidation");

//            boolean validationSuccessful = validateColumns("src/main/resources/excelFiles/file1.xls", "id",
//                    "src/main/resources/excelFiles/file2.xls", "id","null");


            if (validationSuccessful) {
                System.out.println("Columns match.");
            } else {
                System.out.println("Columns do not match.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static boolean validateColumns(String file1Path, String column1Name, String file2Path, String column2Name,
                                          String validationModel) throws IOException {

        try (FileInputStream file1InputStream = new FileInputStream(file1Path);
             FileInputStream file2InputStream = new FileInputStream(file2Path);
             Workbook workbook1 = WorkbookFactory.create(file1InputStream);
             Workbook workbook2 = WorkbookFactory.create(file2InputStream)) {

            Sheet sheet1 = workbook1.getSheetAt(0); // Assuming you want to check the first sheet
            Sheet sheet2 = workbook2.getSheetAt(0);

            int columnIndex1 = -1;
            int columnIndex2 = -1;

            Row headerRow1 = sheet1.getRow(0);
            Row headerRow2 = sheet2.getRow(0);

            // Find the column indices by column name
            for (Cell cell : headerRow1) {
                if (cell.getStringCellValue().trim().equals(column1Name)) {
                    columnIndex1 = cell.getColumnIndex();
                    break;
                }
            }

            for (Cell cell : headerRow2) {
                if (cell.getStringCellValue().trim().equals(column2Name)) {
                    columnIndex2 = cell.getColumnIndex();
                    break;
                }
            }

            if (columnIndex1 == -1 || columnIndex2 == -1) {
                System.out.println("One or both of the specified columns were not found.");
                return false;
            }

            for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                Row row1 = sheet1.getRow(i);
                Row row2 = sheet2.getRow(i);

                Cell cell1 = row1.getCell(columnIndex1);
                Cell cell2 = row2.getCell(columnIndex2);

                String value1 = getCellValueAsString(cell1).trim();
                String value2 = getCellValueAsString(cell2).trim();

                if (!areEquivalent(value1, value2, validationModel)) {
                    System.out.println("Mismatch found at row " + (i + 1) + ":");
                    System.out.println(column1Name + " in File 1: " + value1);
                    System.out.println(column2Name + " in File 2: " + value2);
                    return false;
                }
            }
            return true; // No mismatches found
        }
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

    public static boolean areEquivalent(String value1, String value2, String validationModel) {
        // Check if the values are the same or equivalent based on the input
        if (value1.equals(value2)) {
            return true;
        }

        if (validationModel.equals("specialValidation")) {

            Map<String, String> equivalenceMap = new HashMap<>();
            equivalenceMap.put("Y", "Yes");
            equivalenceMap.put("N", "No");

            String equivalentValue1 = equivalenceMap.get(value1);
            return equivalentValue1 != null && equivalentValue1.equals(value2);
        }

        return false; // No equivalence or exact match
    }
}
