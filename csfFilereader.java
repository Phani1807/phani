public class MapToExcelConverter {

    public static void main(String[] args) {
        // Sample data (replace with your actual Map)
        Map<String, Map<String, Map<String, Set<String>>>> myNestedMap = Map.of(
                "productA", Map.of(
                        "color", Map.of(
                                "red", Set.of("small", "medium"),
                                "blue", Set.of("large")
                        ),
                        "material", Map.of(
                                "cotton", Set.of("available"),
                                "silk", Set.of("out of stock")
                        )
                ),
                "productB", Map.of(
                        "size", Map.of(
                                "s", Set.of("in stock"),
                                "m", Set.of("low stock")
                        )
                )
        );

        String targetProduct = "productA";
        List<List<String>> excelData = fetchDataForProduct(myNestedMap, targetProduct);
        String outputFilePath = "C:\\Users\\inahp\\eclipse-workspace\\ImageToExcel\\src\\main\\resources\\images\\productA_data.xlsx";
        writeToExcel(excelData, outputFilePath);
    }

    public static List<List<String>> fetchDataForProduct(Map<String, Map<String, Map<String, Set<String>>>> nestedMap, String targetProduct) {
        List<List<String>> excelData = new ArrayList<>();

        if (nestedMap.containsKey(targetProduct)) {
            Map<String, Map<String, Set<String>>> productData = nestedMap.get(targetProduct);
            String key1 = targetProduct; // The first key is now fixed

            for (Map.Entry<String, Map<String, Set<String>>> entryLevel2 : productData.entrySet()) {
                String key2 = entryLevel2.getKey();
                for (Map.Entry<String, Set<String>> entryLevel3 : entryLevel2.getValue().entrySet()) {
                    String key3 = entryLevel3.getKey();
                    for (String value : entryLevel3.getValue()) {
                        List<String> rowData = new ArrayList<>();
                        rowData.add(key1);
                        rowData.add(key2);
                        rowData.add(key3);
                        rowData.add(value);
                        excelData.add(rowData);
                    }
                }
            }
        } else {
            System.out.println("Product '" + targetProduct + "' not found in the data.");
        }
        return excelData;
    }

    public static void writeToExcel(List<List<String>> data, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");

            // Create header row (optional)
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Product");
            headerRow.createCell(1).setCellValue("Key Level 2");
            headerRow.createCell(2).setCellValue("Key Level 3");
            headerRow.createCell(3).setCellValue("Value");

            int rowNum = 1;
            for (List<String> rowData : data) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String cellData : rowData) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(cellData);
                }
            }

            // Write the workbook to a file
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
            System.out.println("Data for '" + filePath + "' written to Excel successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
