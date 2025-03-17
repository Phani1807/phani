import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.stream.Collectors;

public class FilterCSVToExcelWithMultipleSheets {

    public static void main(String[] args) {
        String csvFilePath = "input.csv";
        String xlsxFilePath = "output.xlsx";

        try (Reader reader = new BufferedReader(new InputStreamReader(new FileInputStream(csvFilePath), StandardCharsets.UTF_8));
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader());
             Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(xlsxFilePath)) {

            Sheet openSheet = workbook.createSheet("Open Status");
            Sheet inProgressOrReviewSheet = workbook.createSheet("In Progress/Review");

            List<CSVRecord> openRecords = csvParser.getRecords().stream()
                    .filter(record -> "Open".equalsIgnoreCase(record.get("Status")))
                    .collect(Collectors.toList());

            List<CSVRecord> inProgressOrReviewRecords = csvParser.getRecords().stream()
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
