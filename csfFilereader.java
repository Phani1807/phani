import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelSheetSplitterSingleFile {

    public static void main(String[] args) {
        String inputFilePath = "input.xlsx"; // Replace with your input file path
        String outputFilePath = "output.xlsx"; // output file with two sheets

        try (FileInputStream inputStream = new FileInputStream(inputFilePath);
             Workbook inputWorkbook = new XSSFWorkbook(inputStream);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet inputSheet = inputWorkbook.getSheetAt(0); // Assuming data is in the first sheet

            Sheet openSheet = outputWorkbook.createSheet("Open");
            Sheet inReviewInProgressSheet = outputWorkbook.createSheet("In Review/In Progress");

            int openRowCount = 0;
            int inReviewInProgressRowCount = 0;

            // Copy header row
            Row headerRow = inputSheet.getRow(0);
            if (headerRow != null) {
                copyRow(headerRow, openSheet.createRow(openRowCount++));
                copyRow(headerRow, inReviewInProgressSheet.createRow(inReviewInProgressRowCount++));
            }

            // Iterate through data rows
            for (int rowIndex = 1; rowIndex <= inputSheet.getLastRowNum(); rowIndex++) {
                Row dataRow = inputSheet.getRow(rowIndex);
                if (dataRow != null) {
                    Cell statusCell = dataRow.getCell(getColumnIndex(headerRow, "status")); // Assuming "status" column exists.

                    if (statusCell != null) {
                        String status = statusCell.getStringCellValue().trim();

                        if (status.equalsIgnoreCase("open")) {
                            copyRow(dataRow, openSheet.createRow(openRowCount++));
                        } else if (status.equalsIgnoreCase("in review") || status.equalsIgnoreCase("in progress")) {
                            copyRow(dataRow, inReviewInProgressSheet.createRow(inReviewInProgressRowCount++));
                        }
                    }
                }
            }

            // Write the output workbook to a single file
            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(outputStream);
            }

            System.out.println("Excel sheets split successfully into a single file.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void copyRow(Row sourceRow, Row destinationRow) {
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = destinationRow.createCell(i);

            if (oldCell != null) {
                copyCell(oldCell, newCell);
            }
        }
    }

    private static void copyCell(Cell oldCell, Cell newCell) {
        switch (oldCell.getCellType()) {
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(oldCell)) {
                    newCell.setCellValue(oldCell.getDateCellValue());
                } else {
                    newCell.setCellValue(oldCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case FORMULA:
                newCell.setCellFormula(oldCell.getCellFormula());
                break;
            case BLANK:
                break;
            default:
                break;
        }
    }

    private static int getColumnIndex(Row headerRow, String columnName) {
        if (headerRow != null) {
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);
                if (cell != null && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    return i;
                }
            }
        }
        return -1; // Column not found
    }
}
