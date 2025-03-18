import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.stream.Collectors;

public class FilterCSVToExcelWithMultipleSheetsCorrected {

    public static void main(String[] args) {
        String csvFilePath = "input.csv";
        String xlsxFilePath = "output.xlsx";

        try (Reader reader = new BufferedReader(new InputStreamReader(new FileInputStream(csvFilePath), StandardCharsets.UTF_8));
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader());
             Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(xlsxFilePath)) {

            Sheet openSheet = workbook.createSheet("Open Status");
            Sheet inProgressOrReviewSheet = workbook.createSheet("In Progress/Review");

            List<CSVRecord> allRecords = csvParser.getRecords(); // Read all records into a list

            List<CSVRecord> openRecords = allRecords.stream()
                    .filter(record -> "Open".equalsIgnoreCase(record.get("Status")))
                    .collect(Collectors.toList());

            List<CSVRecord> inProgressOrReviewRecords = allRecords.stream()
                    .filter(record -> !"Open".equalsIgnoreCase(record.get("Status")))
                    .collect(Collectors.toList());

            // Write headers to both sheets
            Row openHeaderRow = openSheet.createRow(0);
            Row inProgressOrReviewHeaderRow = inProgressOrReviewSheet.createRow(0);
            String[] headers = csvParser.getHeaderNames().toArray(new String[0]);
            for (int i = 0; i < headers.length; i++) {
                Cell openCell = openHeaderRow.createCell(i);
                openCell.setCellValue(headers[i]);
                Cell inProgressOrReviewCell = inProgressOrReviewHeaderRow.createCell(i);
                inProgressOrReviewCell.setCellValue(headers[i]);
            }

            // Write "Open" records to "Open Status" sheet
            int openRowNum = 1;
            for (CSVRecord record : openRecords) {
                Row row = openSheet.createRow(openRowNum++);
                for (int i = 0; i < record.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(record.get(i));
                }
            }

            // Write "In Progress" and "In Review" records to "In Progress/Review" sheet
            int inProgressOrReviewRowNum = 1;
            for (CSVRecord record : inProgressOrReviewRecords) {
                Row row = inProgressOrReviewSheet.createRow(inProgressOrReviewRowNum++);
                for (int i = 0; i < record.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(record.get(i));
                }
            }

            workbook.write(outputStream);
            System.out.println("Filtered data written to Excel with multiple sheets.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}










import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class AutoSizeExcelColumns {

    public static void main(String[] args) {
        String filePath = "input.xlsx"; // Replace with your file path

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis);
             FileOutputStream fos = new FileOutputStream(filePath)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                autoSizeColumns(sheet);
            }

            workbook.write(fos);
            System.out.println("Columns auto-sized in all sheets.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void autoSizeColumns(Sheet sheet) {
        int lastColumn = 0;
        for (Row row : sheet) {
            if (row != null) {
                lastColumn = Math.max(lastColumn, row.getLastCellNum());
            }
        }
        for (int i = 0; i < lastColumn; i++) {
            sheet.autoSizeColumn(i);
        }
    }
}







import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class UpdateExcelColumnHeaders {

    public static void main(String[] args) {
        String filePath = "input.xlsx"; // Replace with your file path

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = WorkbookFactory.create(fis);
             FileOutputStream fos = new FileOutputStream(filePath)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming the headers are in the first sheet

            Row headerRow = sheet.getRow(0); // Assuming the headers are in the first row

            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        if (cellValue.startsWith("CustomField(") && cellValue.endsWith(")")) {
                            String newCellValue = cellValue.substring("CustomField(".length(), cellValue.length() - 1);
                            cell.setCellValue(newCellValue);
                        }
                    }
                }
            }

            workbook.write(fos);
            System.out.println("Column headers updated.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
