import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class RemoveEmptyRows {

    public static void main(String[] args) {
        String filePath = "path/to/your/excel/file.xlsx"; // Replace with the actual path to your Excel file

        try {
            removeEmptyRows(filePath);
            System.out.println("Successfully removed empty rows from: " + filePath);
        } catch (IOException e) {
            System.err.println("Error processing Excel file: " + e.getMessage());
        }
    }

    public static void removeEmptyRows(String filePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream); // Assuming .xlsx format, use HSSFWorkbook for .xls
        Sheet sheet = workbook.getSheetAt(0); // Assuming you want to process the first sheet

        List<Integer> emptyRowIndices = new ArrayList<>();
        Iterator<Row> rowIterator = sheet.iterator();

        // Identify empty rows
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (isRowEmpty(row)) {
                emptyRowIndices.add(row.getRowNum());
            }
        }

        // Shift rows to remove empty ones (iterating in reverse to avoid index issues)
        for (int i = emptyRowIndices.size() - 1; i >= 0; i--) {
            int rowIndexToDelete = emptyRowIndices.get(i);
            sheet.removeRow(sheet.getRow(rowIndexToDelete));
        }

        // Adjust row numbers for subsequent rows
        int shift = 0;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                shift++;
                continue;
            }
            if (shift > 0) {
                sheet.shiftRows(i, sheet.getLastRowNum(), -shift);
                break; // Only need to trigger the shift once from the first removed block
            }
        }

        // Write the changes back to the Excel file
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
        fileInputStream.close();
    }

    private static boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }
        for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
            Cell cell = row.getCell(c);
            if (cell != null && cell.getCellType() != CellType.BLANK && cell.toString().trim().length() > 0) {
                return false;
            }
        }
        return true;
    }
}
