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
                        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
                        cell.setCellValue((java.util.Date) value);
                        cell.setCellStyle(cellStyle);
                    } else {
                        cell.setCellValue(value.toString());
                    }
