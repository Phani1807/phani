import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class excelmerger12 {
	private static final String[] KEY_COLUMNS = {"feedId","GOC","sub_glc"};
    private static final String[] VALUE_COLUMNS = {"creditamount", "sampleamount"};
    
    public static void main(String[] args) {
        String[] inputFiles = {
        		System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR.xlsx", // Replace with your file paths
        		System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_2.xlsx",
        		System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_3.xlsx",
        		System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_4.xlsx"
        };
        String outputFile = System.getProperty("user.dir") + "\\src\\main\\resources\\images\\output.xlsx";
        
        try {
            mergeExcelFiles(inputFiles, outputFile);
            System.out.println("Grouped comparative merge completed successfully!");
        } catch (IOException e) {
            System.err.println("Error during merge: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    private static void mergeExcelFiles(String[] inputFiles, String outputFile) throws IOException {
        // Use LinkedHashMap to maintain insertion order
        Map<String, Map<String, Map<String, String>>> mergedData = new LinkedHashMap<>();
        // Track row order across files
        final Map<String, Integer> rowOrder = new HashMap<>();
        int orderCounter = 0;
        
        // Process each input file
        for (String inputFile : inputFiles) {
            try (FileInputStream fis = new FileInputStream(inputFile);
                 Workbook workbook = new XSSFWorkbook(fis)) {
                
                Sheet sheet = workbook.getSheetAt(0);
                Map<String, Integer> columnIndices = getColumnIndices(sheet);
                validateColumns(columnIndices, inputFile);
                
                String fileId = inputFile.replaceAll("\\.xlsx$", "");
                
                // Process each row
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) continue; // Skip header row
                    
                    String combinedKey = generateCombinedKey(row, columnIndices);
                    
                    // Track order of first appearance
                    if (!rowOrder.containsKey(combinedKey)) {
                        rowOrder.put(combinedKey, orderCounter++);
                    }
                    
                    // Get or create data map for this key
                    Map<String, Map<String, String>> keyData = mergedData.computeIfAbsent(
                        combinedKey, k -> new LinkedHashMap<>()
                    );
                    
                    // Create map for this file's values
                    Map<String, String> fileValues = new LinkedHashMap<>();
                    
                    // Process value columns
                    for (String valueColumn : VALUE_COLUMNS) {
                        int colIndex = columnIndices.get(valueColumn);
                        Cell cell = row.getCell(colIndex);
                        String value = getCellValueAsString(cell);
                        fileValues.put(valueColumn, value);
                    }
                    
                    keyData.put(fileId, fileValues);
                }
            }
        }
        
        // Create output workbook
        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Sheet outputSheet = outputWorkbook.createSheet("Grouped Comparative Data");
            
            // Create header row
            Row headerRow = outputSheet.createRow(0);
            int colIndex = 0;
            
            // Add key columns to header
            for (String keyColumn : KEY_COLUMNS) {
                headerRow.createCell(colIndex++).setCellValue(keyColumn);
            }
            
            // Add grouped value columns
            for (String valueColumn : VALUE_COLUMNS) {
                // Add all file columns for this value
                for (String inputFile : inputFiles) {
                	 File file = new File(inputFile);
                     String fileName = file.getName();
                    String columnHeader = fileName.substring(0, fileName.lastIndexOf('.')) + "_" + valueColumn;
                    headerRow.createCell(colIndex++).setCellValue(columnHeader);
                }
            }
            
            // Sort entries based on original row order
            List<Map.Entry<String, Map<String, Map<String, String>>>> sortedEntries = 
                new ArrayList<>(mergedData.entrySet());
            sortedEntries.sort(Comparator.comparing(e -> rowOrder.get(e.getKey())));
            
            // Add data rows in sorted order
            int rowIndex = 1;
            for (Map.Entry<String, Map<String, Map<String, String>>> entry : sortedEntries) {
                Row dataRow = outputSheet.createRow(rowIndex++);
                
                // Split combined key back into individual values
                String[] keyValues = entry.getKey().split("\\|");
                colIndex = 0;
                
                // Write key values
                for (String keyValue : keyValues) {
                    dataRow.createCell(colIndex++).setCellValue(keyValue);
                }
                
                // Write grouped values
                Map<String, Map<String, String>> keyData = entry.getValue();
                
                // For each value column
                for (String valueColumn : VALUE_COLUMNS) {
                    // Write values from all files for this column
                    for (String inputFile : inputFiles) {
                        String fileId = inputFile.replaceAll("\\.xlsx$", "");
                        Map<String, String> fileValues = keyData.getOrDefault(fileId, new HashMap<>());
                        String value = fileValues.getOrDefault(valueColumn, "");
                        
                        Cell cell = dataRow.createCell(colIndex++);
                        cell.setCellValue(value);
                    }
                }
            }
            
            // Add group separator borders
            int numColumns = KEY_COLUMNS.length + (VALUE_COLUMNS.length * inputFiles.length);
           // Auto-size columns
            for (int i = 0; i < numColumns; i++) {
                outputSheet.autoSizeColumn(i);
            }
            
            // Write output file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }
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
    
    private static void validateColumns(Map<String, Integer> columnIndices, String fileName) {
        List<String> missingColumns = new ArrayList<>();
        
        // Check key columns
        for (String keyColumn : KEY_COLUMNS) {
            if (!columnIndices.containsKey(keyColumn)) {
                missingColumns.add(keyColumn);
            }
        }
        
        // Check value columns
        for (String valueColumn : VALUE_COLUMNS) {
            if (!columnIndices.containsKey(valueColumn)) {
                missingColumns.add(valueColumn);
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