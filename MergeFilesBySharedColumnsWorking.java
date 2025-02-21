import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class MergeFilesBySharedColumnsWorking {

	public static void main(String[] args) {

		String file1Column10 = "creditamount"; // Column to add from file1
		String file2Column12 = "sampleamount"; // Column to add from file2
		String file1Path = System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_3.xlsx";

		String file2Path = System.getProperty("user.dir") + "\\src\\main\\resources\\images\\FEED_MSTR_4.xlsx";

		String[] sharedColumnNames = { "feedId", "GOC", "sub_glc" }; // Replace with your actual keys
		String outputFilePath = System.getProperty("user.dir")+ "\\src\\main\\resources\\images\\merged_two_files.xlsx";
		try (Workbook outputWorkbook = new XSSFWorkbook();
				FileInputStream file1InputStream = new FileInputStream(file1Path);
				FileInputStream file2InputStream = new FileInputStream(file2Path);
				Workbook file1Workbook = WorkbookFactory.create(file1InputStream);
				Workbook file2Workbook = WorkbookFactory.create(file2InputStream)) {

			Sheet outputSheet = outputWorkbook.createSheet("MergedData");
			Sheet file1Sheet = file1Workbook.getSheetAt(0);
			Sheet file2Sheet = file2Workbook.getSheetAt(0);

			// Create header row in output
			Row outputHeaderRow = outputSheet.createRow(0);
			int outputColumnIndex = 0;
			for (String columnName : sharedColumnNames) {
				outputHeaderRow.createCell(outputColumnIndex++).setCellValue(columnName);
			}
			outputHeaderRow.createCell(outputColumnIndex++).setCellValue("file1_" + file1Column10);
			outputHeaderRow.createCell(outputColumnIndex).setCellValue("file2_" + file2Column12);

			// Find column indices
			Map<String, Integer> file1ColumnIndices = getColumnIndices(file1Sheet);
			Map<String, Integer> file2ColumnIndices = getColumnIndices(file2Sheet);

			int file1Column10Index = file1ColumnIndices.getOrDefault(file1Column10, -1);
			int file2Column12Index = file2ColumnIndices.getOrDefault(file2Column12, -1);

			// Build a matching key for each row in file1.
			for (int file1RowIndex = 1; file1RowIndex <= file1Sheet.getLastRowNum(); file1RowIndex++) {
				Row file1Row = file1Sheet.getRow(file1RowIndex);
				if (file1Row == null)
					continue; // Skip empty rows

				// Build a combined key from shared columns
				StringBuilder file1KeyBuilder = new StringBuilder();
				for (String sharedColumnName : sharedColumnNames) {
					int sharedColumnIndex = file1ColumnIndices.getOrDefault(sharedColumnName, -1);
					if (sharedColumnIndex != -1 && file1Row.getCell(sharedColumnIndex) != null) {
						file1KeyBuilder.append(getCellValueAsString(file1Row.getCell(sharedColumnIndex))).append("|"); // using
																														// pipe
																														// as
																														// a
																														// separator.
					}
				}
				String file1Key = file1KeyBuilder.toString();

				// Find matching row in file2
				for (int file2RowIndex = 1; file2RowIndex <= file2Sheet.getLastRowNum(); file2RowIndex++) {
					Row file2Row = file2Sheet.getRow(file2RowIndex);
					if (file2Row == null)
						continue;

					StringBuilder file2KeyBuilder = new StringBuilder();
					for (String sharedColumnName : sharedColumnNames) {
						int sharedColumnIndex = file2ColumnIndices.getOrDefault(sharedColumnName, -1);
						if (sharedColumnIndex != -1 && file2Row.getCell(sharedColumnIndex) != null) {
							file2KeyBuilder.append(getCellValueAsString(file2Row.getCell(sharedColumnIndex)))
									.append("|");
						}
					}
					String file2Key = file2KeyBuilder.toString();

					if (file1Key.equals(file2Key)) {
						Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
						outputColumnIndex = 0;

						// Add shared columns
						for (String sharedColumnName : sharedColumnNames) {
							int sharedColumnIndex = file1ColumnIndices.getOrDefault(sharedColumnName, -1);
							if (sharedColumnIndex != -1 && file1Row.getCell(sharedColumnIndex) != null) {
								copyCellValue(file1Row.getCell(sharedColumnIndex),
										outputRow.createCell(outputColumnIndex++));
							} else {
								outputRow.createCell(outputColumnIndex++).setCellValue(""); // Handle missing data
							}
						}

						// Add file1 column10
						if (file1Column10Index != -1 && file1Row.getCell(file1Column10Index) != null) {
							copyCellValue(file1Row.getCell(file1Column10Index),
									outputRow.createCell(outputColumnIndex++));
						} else {
							outputRow.createCell(outputColumnIndex++).setCellValue("");
						}

						// Add file2 column12
						if (file2Column12Index != -1 && file2Row.getCell(file2Column12Index) != null) {
							copyCellValue(file2Row.getCell(file2Column12Index),
									outputRow.createCell(outputColumnIndex));
						} else {
							outputRow.createCell(outputColumnIndex).setCellValue("");
						}
						break; // Move to the next row in file1
					}
				}
			}

			try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath)) {
				outputWorkbook.write(fileOutputStream);
			}
			System.out.println("Merged Excel file created successfully.");

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static Map<String, Integer> getColumnIndices(Sheet sheet) {
		Map<String, Integer> columnIndices = new HashMap<String, Integer>();
		Row headerRow = sheet.getRow(0);
		if (headerRow != null) {
			for (Cell cell : headerRow) {
				if (cell.getCellType() == CellType.STRING) {
					columnIndices.put(cell.getStringCellValue(), cell.getColumnIndex());
				}
			}
		}
		return columnIndices;
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
			} else {
				return String.valueOf(cell.getNumericCellValue());
			}
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		default:
			return "";
		}
	}

	private static void copyCellValue(Cell sourceCell, Cell targetCell) {
		switch (sourceCell.getCellType()) {
		case STRING:
			targetCell.setCellValue(sourceCell.getStringCellValue());
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(sourceCell)) {
				targetCell.setCellValue(sourceCell.getDateCellValue());
			} else {
				targetCell.setCellValue(sourceCell.getNumericCellValue());
			}
			break;
		case BOOLEAN:
			targetCell.setCellValue(sourceCell.getBooleanCellValue());
			break;
		case FORMULA:
			targetCell.setCellFormula(sourceCell.getCellFormula());
			break;
		case BLANK:
			targetCell.setBlank();
			break;
		default:
			break;
		}
	}
}