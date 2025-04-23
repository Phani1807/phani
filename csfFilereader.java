String currentProduct = null;
            int productMergeStart = -1;
            String currentKeyLevel2 = null;
            int keyLevel2MergeStart = -1;

            for (int i = 0; i < data.size(); i++) {
                List<String> rowData = data.get(i);
                Row row = sheet.createRow(rowNum++);
                String product = rowData.get(0);
                String keyLevel2 = rowData.get(1);
                String keyLevel3 = rowData.get(2);
                String value = rowData.get(3);

                row.createCell(0).setCellValue(product);
                row.createCell(1).setCellValue(keyLevel2);
                row.createCell(2).setCellValue(keyLevel3);
                row.createCell(3).setCellValue(value);

                // Merge cells for "Product"
                if (!product.equals(currentProduct)) {
                    if (productMergeStart != -1 && i > productMergeStart) {
                        sheet.addMergedRegion(new CellRangeAddress(productMergeStart + 1, i, 0, 0));
                    }
                    currentProduct = product;
                    productMergeStart = i;
                    keyLevel2MergeStart = i; // Reset Key Level 2 merge start when Product changes
                    currentKeyLevel2 = keyLevel2;
                }

                // Merge cells for "Key Level 2"
                if (product.equals(currentProduct)) {
                    if (!keyLevel2.equals(currentKeyLevel2)) {
                        if (keyLevel2MergeStart != -1 && i > keyLevel2MergeStart) {
                            sheet.addMergedRegion(new CellRangeAddress(keyLevel2MergeStart + 1, i, 1, 1));
                        }
                        currentKeyLevel2 = keyLevel2;
                        keyLevel2MergeStart = i;
                    }
                }

                // Handle the last row for merging
                if (i == data.size() - 1) {
                    if (productMergeStart != -1 && i >= productMergeStart) {
                        sheet.addMergedRegion(new CellRangeAddress(productMergeStart + 1, i + 1, 0, 0));
                    }
                    if (keyLevel2MergeStart != -1 && i >= keyLevel2MergeStart) {
                        sheet.addMergedRegion(new CellRangeAddress(keyLevel2MergeStart + 1, i + 1, 1, 1));
                    }
                }
            }
