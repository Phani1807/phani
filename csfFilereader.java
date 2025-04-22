import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.IntStream;

public class ExcelFormatterMerged {

    public static void main(String[] args) {
        String filePath = "path/to/your/excel/file.xlsx";
        try {
            formatExcel(filePath);
            System.out.println("Successfully removed empty rows and autosized columns in: " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void formatExcel(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook wb = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(filePath)) {
            Sheet sh = wb.getSheetAt(0);

            // Remove empty rows
            List<Integer> rowsToRemove = new ArrayList<>();
            Iterator<Row> rowIterator = sh.iterator();
            int rowIndex = 0;
            while (rowIterator.hasNext()) {
                if (isRowEmpty(rowIterator.next())) {
                    rowsToRemove.add(rowIndex);
                }
                rowIndex++;
            }
            for (int i = rowsToRemove.size() - 1; i >= 0; i--) {
                sh.removeRow(sh.getRow(rowsToRemove.get(i)));
            }

            // Autosize columns (after removing empty rows to get accurate last row)
            if (sh.getRow(0) != null) { // Check if the sheet has any rows
                IntStream.range(0, sh.getRow(0).getLastCellNum()).forEach(sh::autoSizeColumn);
            }

            wb.write(fos);
        }
    }

    private static boolean isRowEmpty(Row row) {
        return row == null || IntStream.rangeClosed(row.getFirstCellNum(), row.getLastCellNum())
                                        .mapToObj(row::getCell)
                                        .allMatch(cell -> cell == null || cell.getCellType() == CellType.BLANK || cell.toString().trim().isEmpty());
    }
}
