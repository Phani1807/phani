package feb24;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelMergerCustomColumns {
    private static final String[] KEY_COLUMNS = {"feedId", "GOC", "sub_glc"};
    
    public static void main(String[] args) {
        String[] inputFiles = {
            System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR.xlsx",
            System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_2.xlsx",
            System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_3.xlsx",
            System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_4.xlsx"
        };
        String outputFile = System.getProperty("user.dir") + "\\src\\main\\resources\\images\\output2.xlsx";
        
        // Define user-requested columns as a map of fileIndex to columnName
        // Map format: fileIndex (0-based) -> List of column names to include from that file
        Map<Integer, List<String>> requestedColumns = new HashMap<>();
        requestedColumns.put(0, Arrays.asList("creditamount","sampleamount"));       // From file 1
        requestedColumns.put(1, Arrays.asList("sampleamount"));       // From file 3
        requestedColumns.put(3, Arrays.asList("creditamount"));       // From file 4
        
        try {
            mergeExcelFiles(inputFiles, outputFile, requestedColumns);
            System.out.println("Custom column merge completed successfully!");
        } catch (IOException e) {
            System.err.println("Error during merge: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static void mergeExcelFiles(String[] inputFiles, String outputFile, 
                                       Map<Integer, List<String>> requestedColumns) throws IOException {
        // Use LinkedHashMap to maintain insertion order
        Map<String, Map<Integer, Map<String, String>>> mergedData = new LinkedHashMap<>();
        // Track row order across files
        final Map<String, Integer> rowOrder = new HashMap<>();
        int orderCounter = 0;
        
        // Process each input file
        for (int fileIndex = 0; fileIndex < inputFiles.length; fileIndex++) {
            String inputFile = inputFiles[fileIndex];
            
            // Skip files with no requested columns
            if (!requestedColumns.containsKey(fileIndex)) {
                continue;
            }
            
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = new XSSFWorkbook(fis)) {
                
                Sheet sheet = workbook.getSheetAt(0);
                Map<String, Integer> columnIndices = getColumnIndices(sheet);
                validateColumns(columnIndices, inputFile, requestedColumns.get(fileIndex));
                
                // Process each row
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) continue; // Skip header row
                    
                    String combinedKey = generateCombinedKey(row, columnIndices);
                    
                    // Track order of first appearance
                    if (!rowOrder.containsKey(combinedKey)) {
                        rowOrder.put(combinedKey, orderCounter++);
                    }
                    
                    // Get or create data map for this key
                    Map<Integer, Map<String, String>> keyData = mergedData.computeIfAbsent(
                        combinedKey, k -> new LinkedHashMap<>()
                    );
                    
                    // Create map for this file's values
                    Map<String, String> fileValues = new LinkedHashMap<>();
                    
                    // Process requested columns for this file
                    for (String columnName : requestedColumns.get(fileIndex)) {
                        int colIndex = columnIndices.get(columnName);
                        Cell cell = row.getCell(colIndex);
                        String value = getCellValueAsString(cell);
                        fileValues.put(columnName, value);
                    }
                    
                    keyData.put(fileIndex, fileValues);
                }
            }
        }
        
        // Create output workbook
        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Sheet outputSheet = outputWorkbook.createSheet("Custom Column Data");
            
            // Create header row
            Row headerRow = outputSheet.createRow(0);
            int colIndex = 0;
            
            // Add key columns to header
            for (String keyColumn : KEY_COLUMNS) {
                headerRow.createCell(colIndex++).setCellValue(keyColumn);
            }
            
            // Create a list of all output columns for easy reference
            List<OutputColumn> outputColumns = new ArrayList<>();
            
            // Add user-requested columns to header
            for (int fileIndex : requestedColumns.keySet()) {
                String fileName = new File(inputFiles[fileIndex]).getName();
                fileName = fileName.substring(0, fileName.lastIndexOf('.'));
                
                for (String columnName : requestedColumns.get(fileIndex)) {
                    String columnHeader = fileName + "_" + columnName;
                    headerRow.createCell(colIndex++).setCellValue(columnHeader);
                    
                    // Add to tracking list
                    outputColumns.add(new OutputColumn(fileIndex, columnName));
                }
            }
            
            // Sort entries based on original row order
            List<Map.Entry<String, Map<Integer, Map<String, String>>>> sortedEntries = 
                new ArrayList<>(mergedData.entrySet());
            sortedEntries.sort(Comparator.comparing(e -> rowOrder.get(e.getKey())));
            
            // Add data rows in sorted order
            int rowIndex = 1;
            for (Map.Entry<String, Map<Integer, Map<String, String>>> entry : sortedEntries) {
                Row dataRow = outputSheet.createRow(rowIndex++);
                
                // Split combined key back into individual values
                String[] keyValues = entry.getKey().split("\\|");
                colIndex = 0;
                
                // Write key values
                for (String keyValue : keyValues) {
                    dataRow.createCell(colIndex++).setCellValue(keyValue);
                }
                
                // Write requested columns
                Map<Integer, Map<String, String>> keyData = entry.getValue();
                
                // For each output column
                for (OutputColumn outputColumn : outputColumns) {
                    Map<String, String> fileValues = keyData.getOrDefault(outputColumn.fileIndex, new HashMap<>());
                    String value = fileValues.getOrDefault(outputColumn.columnName, "");
                    
                    Cell cell = dataRow.createCell(colIndex++);
                    cell.setCellValue(value);
                }
            }
            
            // Auto-size columns
            for (int i = 0; i < colIndex; i++) {
                outputSheet.autoSizeColumn(i);
            }
            
            // Write output file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
        }
    }
    
    // Helper class to track output columns
    private static class OutputColumn {
        final int fileIndex;
        final String columnName;
        
        OutputColumn(int fileIndex, String columnName) {
            this.fileIndex = fileIndex;
            this.columnName = columnName;
        }
    }
    
    private static Map<String, Integer> getColumnIndices(Sheet sheet) {
        Map<String, Integer> indices = new HashMap<>();
        Row headerRow = sheet.getRow(0);
        
        for (Cell cell : headerRow) {
            String columnName = cell.getStringCellValue();
            indices.put(columnName, cell.getColumnIndex());
        }
        
        return indices;
    }
    
    private static void validateColumns(Map<String, Integer> columnIndices, String fileName, 
                                       List<String> requestedColumns) {
        List<String> missingColumns = new ArrayList<>();
        
        // Check key columns
        for (String keyColumn : KEY_COLUMNS) {
            if (!columnIndices.containsKey(keyColumn)) {
                missingColumns.add(keyColumn);
            }
        }
        
        // Check requested columns
        for (String columnName : requestedColumns) {
            if (!columnIndices.containsKey(columnName)) {
                missingColumns.add(columnName);
            }
        }
        
        if (!missingColumns.isEmpty()) {
            throw new IllegalStateException("Missing required columns in " + fileName + ": " + 
                                         String.join(", ", missingColumns));
        }
    }
    
    private static String generateCombinedKey(Row row, Map<String, Integer> columnIndices) {
        StringBuilder key = new StringBuilder();
        
        for (String keyColumn : KEY_COLUMNS) {
            if (key.length() > 0) {
                key.append("|");
            }
            Cell cell = row.getCell(columnIndices.get(keyColumn));
            key.append(getCellValueAsString(cell));
        }
        
        return key.toString();
    }
    
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
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