// If there are are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                        )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }

    private void readColumnHeader(XSSFSheet sheetFile) {
		if (columnHeaders == null) {
			cellStyles = new CellStyle[NUM_EXCEL_COLUMNS];
			columnHeaders = new String[NUM_EXCEL_COLUMNS];
			origRow = sheetFile.getRow(HEADER_ROW_INDEX);
		}
		Row row = sheetFile.getRow(HEADER_ROW_INDEX);
		for (int i = START_COLUMN_INDEX, myIndex = 0; i < (START_COLUMN_INDEX + NUM_EXCEL_COLUMNS); i++, myIndex++) {
			cellStyles[myIndex] = row.getCell(i).getCellStyle();
			columnHeaders[myIndex] = row.getCell(i).getStringCellValue();
		}
	}

	private void writeColumnHeader(Workbook workbook, XSSFSheet sheetFile) {
		Row row = sheetFile.createRow(HEADER_ROW_INDEX);
		row = origRow;
		CellStyle newCellStyle;
		Cell cell;
		for (int i = START_COLUMN_INDEX, myIndex = 0; i < columnHeaders.length; i++, myIndex++) {
			cell = row.createCell(i);
			newCellStyle =  workbook.createCellStyle();
			newCellStyle.cloneStyleFrom(cellStyles[myIndex]);
			cell.setCellStyle(newCellStyle);
			cell.setCellValue(columnHeaders[myIndex]);
		}
	}
	
	private static void copyRow(XSSFWorkbook newWorkbook, XSSFSheet origWorksheet, XSSFSheet newWorksheet, int sourceRowNum, int destinationRowNum, boolean isHeader) {
        // Get the source / new row
        Row newRow = newWorksheet.getRow(destinationRowNum);
        Row sourceRow = origWorksheet.getRow(sourceRowNum);
        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            newWorksheet.shiftRows(destinationRowNum, newWorksheet.getLastRowNum(), 1);
        } else {
            newRow = newWorksheet.createRow(destinationRowNum);
        }
        System.out.println("row Height: " + sourceRow.getHeight());
        System.out.println("source row height: " + newRow.getHeight());
        newRow.setHeight(sourceRow.getHeight());
        System.out.println("source row height: " + newRow.getHeight());

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy column widt
            if (isHeader) {
            	newWorksheet.setColumnWidth(i, origWorksheet.getColumnWidth(i));
            }
            // Copy style from old cell and apply to new cell
            CellStyle newCellStyle = newWorkbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            //newCellStyle.setFillBackgroundColor(oldCell.getCellStyle().getFillBackgroundColor());
            //newCellStyle.setFillBackgroundColor((short)200);
            newCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            newCellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
            System.out.println("color: " + oldCell.getCellStyle().getFillBackgroundColor());
            System.out.println("rgb " + oldCell.getCellStyle().getFillBackgroundColorColor());
            System.out.println("pro " + oldCell.getCellStyle().getFillForegroundColor());
            
            CellStyle testCellStyle = oldCell.getCellStyle();
            System.out.println(testCellStyle.getFillBackgroundColor());
            System.out.println(testCellStyle.getFillBackgroundColorColor());
            System.out.println(testCellStyle.getFillForegroundColor());
            System.out.println(testCellStyle.getFillForegroundColorColor());
            System.out.println(testCellStyle.getFillPattern());
            
            newCell.setCellStyle(newCellStyle);
            
            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }
    }
	
	private static void copyRow(XSSFWorkbook newWorkbook, XSSFSheet origWorksheet, XSSFSheet newWorksheet, int sourceRowNum) {
		copyRow(newWorkbook, origWorksheet, newWorksheet, sourceRowNum, sourceRowNum, true);
	}