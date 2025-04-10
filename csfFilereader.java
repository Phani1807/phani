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








import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

public class FilloJiraMigration {

    public static void main(String[] args) {
        String excelFilePath = "path/to/your/jira_data.xlsx"; // Replace with your Excel file path
        Fillo fillo = new Fillo();
        Connection connection = null;

        try {
            connection = fillo.getConnection(excelFilePath);

            String issueKeyToUpdate = "hello"; // The issue key you're working with

            // 1. Update (Optional, if you have any update logic)
            // Example: Update a specific column based on the issue key
            String updateQuery = "UPDATE openjiras SET status='In Progress' WHERE issuekey='" + issueKeyToUpdate + "'";
            connection.executeUpdate(updateQuery);
            System.out.println("Update successful (if applicable).");

            // 2. Select data from openjiras
            String selectQuery = "SELECT * FROM openjiras WHERE issuekey='" + issueKeyToUpdate + "'";
            Recordset recordset = connection.executeQuery(selectQuery);

            if (recordset.next()) { // Check if a record was found
                // Build the insert query
                StringBuilder insertQueryBuilder = new StringBuilder("INSERT INTO progressjiras (");
                StringBuilder valuesBuilder = new StringBuilder("VALUES (");

                // Get column names from the recordset
                for (int i = 0; i < recordset.getFieldNames().size(); i++) {
                    insertQueryBuilder.append(recordset.getFieldNames().get(i));
                    valuesBuilder.append("'").append(recordset.getField(recordset.getFieldNames().get(i))).append("'");

                    if (i < recordset.getFieldNames().size() - 1) {
                        insertQueryBuilder.append(", ");
                        valuesBuilder.append(", ");
                    }
                }

                insertQueryBuilder.append(") ");
                valuesBuilder.append(")");

                // Combine the insert query
                String insertQuery = insertQueryBuilder.toString() + valuesBuilder.toString();
                connection.executeUpdate(insertQuery);
                System.out.println("Insert into progressjiras successful.");

                // Optionally, delete the row from openjiras if needed
                String deleteQuery = "DELETE FROM openjiras WHERE issuekey='" + issueKeyToUpdate + "'";
                connection.executeUpdate(deleteQuery);
                System.out.println("Delete from openjiras successful.");

            } else {
                System.out.println("No record found with issuekey: " + issueKeyToUpdate);
            }

            recordset.close();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (connection != null) {
                connection.close();
            }
        }
    }
}






package april04;
import java.util.Properties;
import javax.mail.*;
import javax.mail.internet.*;

public class OutlookSmtpExample {

    public static void main(String[] args) {

        final String username = "your_outlook_email@outlook.com"; // Replace with your Outlook email
        final String password = "your_outlook_password"; // Replace with your Outlook password or app password

        String host = "smtp-mail.outlook.com"; // Outlook SMTP server
        int port = 587; // Outlook SMTP port

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true"); // Enable STARTTLS
        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", port);

        Session session = Session.getInstance(props,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });

        try {

            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(username));
            message.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse("recipient_email@example.com")); // Replace with recipient
            message.setSubject("Testing Outlook SMTP");
            message.setText("Hello, this is a test email sent using Outlook SMTP.");

            Transport.send(message);

            System.out.println("Email sent successfully!");

        } catch (MessagingException e) {
            e.printStackTrace();
        }
    }
}














import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CsvToXlsxConverter {

    public static void main(String[] args) {
        String csvFilePath = "input.csv"; // Replace with your CSV file path
        String xlsxFilePath = "output.xlsx"; // Replace with your desired XLSX file path

        try {
            convertCsvToXlsx(csvFilePath, xlsxFilePath);
            System.out.println("CSV to XLSX conversion successful!");
        } catch (IOException e) {
            System.err.println("Error converting CSV to XLSX: " + e.getMessage());
            e.printStackTrace();
        }
    }

    public static void convertCsvToXlsx(String csvFilePath, String xlsxFilePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        try (BufferedReader br = new BufferedReader(new FileReader(csvFilePath))) {
            String line;
            int rowNum = 0;
            while ((line = br.readLine()) != null) {
                String[] values = line.split(","); // Assuming comma-separated values

                Row row = sheet.createRow(rowNum++);
                for (int colNum = 0; colNum < values.length; colNum++) {
                    Cell cell = row.createCell(colNum);
                    cell.setCellValue(values[colNum]);
                }
            }
        }

        try (FileOutputStream fos = new FileOutputStream(xlsxFilePath)) {
            workbook.write(fos);
        }
        workbook.close(); // important to close the workbook to release resources.
    }
}
