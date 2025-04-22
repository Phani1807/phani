private static boolean isRowEmpty(Row row) {
    if (row == null) {
        return true;
    }
    for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
        Cell cell = row.getCell(i);
        if (cell != null && cell.getCellType() != CellType.BLANK) {
            String cellValue = cell.toString().trim();
            if (!cellValue.isEmpty()) {
                return false;
            }
        }
    }
    return true;
}
