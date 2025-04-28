import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ExcelWriter {

    public static void writeToExcel(List<Map<String, Object>> data, String filePath) throws IOException {
        // Create a new Excel workbook
        Workbook workbook = new XSSFWorkbook();
        // Create a new sheet
        Sheet sheet = workbook.createSheet("Sheet1");

        int rowNum = 0;
        // Write the header row
        if (!data.isEmpty()) {
            Map<String, Object> firstRow = data.get(0);
            Set<String> headers = firstRow.keySet();
            Row headerRow = sheet.createRow(rowNum++);
            int colNum = 0;
            for (String header : headers) {
                Cell cell = headerRow.createCell(colNum++);
                cell.setCellValue(header);
            }

            // Write the data rows
            for (Map<String, Object> rowData : data) {
                Row dataRow = sheet.createRow(rowNum++);
                colNum = 0;
                for (String header : headers) {
                    Cell cell = dataRow.createCell(colNum++);
                    Object value = rowData.get(header);
                    if (value != null) {
                        if (value instanceof String) {
                            cell.setCellValue((String) value);
                        } else if (value instanceof Integer) {
                            cell.setCellValue((Integer) value);
                        } else if (value instanceof Double) {
                            cell.setCellValue((Double) value);
                        } else if (value instanceof Boolean) {
                            cell.setCellValue((Boolean) value);
                        } else if (value instanceof java.util.Date) {
                            CellStyle cellStyle = workbook.createCellStyle();
                            CreationHelper createHelper = workbook.getCreationHelper();
                            cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss")); // Adjust format as needed
                            cell.setCellValue((java.util.Date) value);
                            cell.setCellStyle(cellStyle);
                        } else {
                            // Handle other data types as needed
                            cell.setCellValue(value.toString());
                        }
                    }
                }
            }

            // Auto-resize columns for better readability
            for (int i = 0; i < headers.size(); i++) {
                sheet.autoSizeColumn(i);
            }
        }

        // Write the workbook to a file
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        }
        workbook.close();

        System.out.println("Data successfully written to: " + filePath);
    }

    public static void main(String[] args) {
        // Example usage:
        List<Map<String, Object>> myData = List.of(
                Map.of("Name", "Alice", "Age", 30, "City", "New York", "Enrollment Date", new java.util.Date()),
                Map.of("Name", "Bob", "Age", 25, "City", "Los Angeles", "Score", 85.5),
                Map.of("Name", "Charlie", "Age", 35, "City", "Chicago", "IsActive", true)
        );

        String excelFilePath = "output.xlsx";

        try {
            writeToExcel(myData, excelFilePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
