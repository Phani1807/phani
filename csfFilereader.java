
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MergeDuplicateColumns {

    public static void main(String[] args) {
    	String filePath = "C:\\Users\\inahp\\eclipse-workspace\\ImageToExcel\\src\\main\\resources\\images\\FEED_MSTR.xlsx"; // Replace with your input file path
        String outputFilePath = "C:\\Users\\inahp\\eclipse-workspace\\ImageToExcel\\src\\main\\resources\\images\\output.xlsx"; // Replace with your output file path


        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            if (sheet != null) {
                mergeDuplicateColumns(sheet);

                try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                    workbook.write(fos);
                }
                System.out.println("Duplicate columns merged and saved to: " + outputFilePath);
            } else {
                System.out.println("Sheet not found.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void mergeDuplicateColumns(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) {
            return; // No header row, nothing to merge
        }

        Map<String, List<Integer>> columnIndices = new HashMap<>();

        // Identify duplicate column names and their indices
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String columnName = cell.getStringCellValue().trim();
                columnIndices.computeIfAbsent(columnName, k -> new ArrayList<>()).add(i);
            }
        }

        // Merge duplicate columns
        for (Map.Entry<String, List<Integer>> entry : columnIndices.entrySet()) {
            List<Integer> indices = entry.getValue();
            if (indices.size() > 1) {
                int targetColumnIndex = indices.get(0); // The first column will be the target

                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        StringBuilder mergedValues = new StringBuilder();
                        for (int columnIndex : indices) {
                            Cell cell = row.getCell(columnIndex);
                            if (cell != null && cell.getCellType() != CellType.BLANK) {

                                String value = "";
                                if(cell.getCellType() == CellType.STRING){
                                    value = cell.getStringCellValue();
                                }else if (cell.getCellType() == CellType.NUMERIC){
                                    value = String.valueOf(cell.getNumericCellValue());
                                }else if (cell.getCellType() == CellType.BOOLEAN){
                                    value = String.valueOf(cell.getBooleanCellValue());
                                }

                                if (!value.trim().isEmpty()) {
                                    if (mergedValues.length() > 0) {
                                        mergedValues.append(", ");
                                    }
                                    mergedValues.append(value.trim());
                                }
                            }
                        }
                        // Write the merged value to the target column
                        Cell targetCell = row.getCell(targetColumnIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        targetCell.setCellValue(mergedValues.toString());

                        // Clear the other duplicate columns
                        for (int i = 1; i < indices.size(); i++) {
                            Cell cellToRemove = row.getCell(indices.get(i), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            cellToRemove.setBlank();
                        }
                    }
                }
                //Remove the header from the extra columns.
                for (int i = 1; i < indices.size(); i++) {
                    Cell cellToRemove = headerRow.getCell(indices.get(i), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellToRemove.setBlank();
                }

            }
        }
    }
}
