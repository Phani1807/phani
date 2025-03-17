import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.stream.Collectors;

public class FilterCSVToExcelWithHeaders {

    public static void main(String[] args) {
        String csvFilePath = "input.csv"; // Replace with your CSV file path
        String xlsxFilePath = "output.xlsx"; // Replace with your XLSX file path

        try (Reader reader = new BufferedReader(new InputStreamReader(new FileInputStream(csvFilePath), StandardCharsets.UTF_8));
             CSVParser csvParser = new CSVParser(reader, CSVFormat.DEFAULT.withFirstRecordAsHeader()); // Read headers from the first record
             Workbook workbook = new XSSFWorkbook();
             FileOutputStream outputStream = new FileOutputStream(xlsxFilePath)) {

            Sheet openSheet = workbook.createSheet("Open Status");

            List<CSVRecord> openRecords = csvParser.getRecords().stream()
                    .filter(record -> "Open".equalsIgnoreCase(record.get("Status")))
                    .collect(Collectors.toList());

            // Write header to the new sheet
            Row headerRow = openSheet.createRow(0);
            String[] headers = csvParser.getHeaderNames().toArray(new String[0]); // Get headers from CSVParser
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Write data to the new sheet
            int rowNum = 1; // Start from row 1 (after header)
            for (CSVRecord record : openRecords) {
                Row row = openSheet.createRow(rowNum++);
                for (int i = 0; i < record.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(record.get(i));
                }
            }

            workbook.write(outputStream);
            System.out.println("Filtered data written to Excel with headers.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}