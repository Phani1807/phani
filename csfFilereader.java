
public static void ExcelDeleteEmptyRows {
		String excelFilePath = "C:\\Users\\inahp\\eclipse-workspace\\ImageToExcel\\src\\main\\resources\\images\\productA_data.xlsx"; // Replace

		try {
			FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = new XSSFWorkbook(fileInputStream);
			Sheet sheet = workbook.getSheet("ActiveJiras"); 

			List<Integer> emptyRows = new ArrayList<>();
			int firstRowNum = sheet.getFirstRowNum();
			int lastRowNum = sheet.getLastRowNum();

			// 1. Find Empty Rows
			for (int i = firstRowNum; i <= lastRowNum; i++) {
				Row row = sheet.getRow(i);
				if (row != null && isRowEmptyscrub(row)) { 
					emptyRows.add(i);
				}
			}

			// 2. Delete Empty Rows (from bottom to top) - Improved Deletion
			for (int i = emptyRows.size() - 1; i >= 0; i--) {
				int rowToDelete = emptyRows.get(i);
				sheet.removeRow(sheet.getRow(rowToDelete)); // Remove the row
			}

			// 3. create a new sheet and copy the non-empty rows
			Sheet newSheet = workbook.createSheet("NewSheet");
			int newRowNum = 0;
			for (int i = firstRowNum; i <= lastRowNum; i++) {
				Row row = sheet.getRow(i);
				if (row != null && !isRowEmptyscrub(row)) {
					Row newRow = newSheet.createRow(newRowNum++);
					copyRowscrub(newSheet, row, newRow);
				}
			}
			// 4. remove the old sheet and rename the new sheet
			workbook.removeSheetAt(workbook.getSheetIndex(sheet));
			workbook.setSheetName(workbook.getSheetIndex(newSheet), "ActiveJiras");

			// 3. Save Changes
			FileOutputStream fileOutputStream = new FileOutputStream(new File(excelFilePath));
			workbook.write(fileOutputStream);

			// Close resources
			fileInputStream.close();
			fileOutputStream.close();
			workbook.close();

			System.out.println("Empty rows deleted and cells shifted up successfully.");

		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("Error processing Excel file: " + e.getMessage());
		}
	
}
	// Helper method to check if a row is empty
	private static boolean isRowEmptyscrub(Row row) {
		if (row == null) {
			return true;
		}
		boolean isEmpty = true;
		for (Cell cell : row) {
			if (cell.getCellType() != CellType.BLANK) {
				isEmpty = false;
				break; // No need to check further if a non-blank cell is found
			}
		}
		return isEmpty;
	}

	private static void copyRowscrub(Sheet sheet, Row sourceRow, Row newRow) {
		for (Cell sourceCell : sourceRow) {
			Cell newCell = newRow.createCell(sourceCell.getColumnIndex(), sourceCell.getCellType()); // create cell with
																										// original type
			copyCellscrub(sourceCell, newCell);
		}
	}

	@SuppressWarnings("deprecation")
	private static void copyCellscrub(Cell oldCell, Cell newCell) {
		CellType cellType = oldCell.getCellType();
		newCell.setCellType(cellType);
		switch (cellType) {
		case BLANK:
			break;
		case BOOLEAN:
			newCell.setCellValue(oldCell.getBooleanCellValue());
			break;
		case ERROR:
			newCell.setCellErrorValue(oldCell.getErrorCellValue());
			break;
		case NUMERIC:
			newCell.setCellValue(oldCell.getNumericCellValue());
			break;
		case STRING:
			newCell.setCellValue(oldCell.getStringCellValue());
			break;
		case FORMULA:
			newCell.setCellFormula(oldCell.getCellFormula());
			break;
		default:
			break;
		}
	}
