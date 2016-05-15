package org.kotemaru.tool.mstotxt;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel implements Processor {

	public void start(String fileName, File file) throws Exception {
		Workbook book = loadBook(file);
		for (int i = 0; i < book.getNumberOfSheets(); i++) {
			Sheet sheet = book.getSheetAt(i);
			procSheet(fileName, sheet);
		}
	}

	private void procSheet(String file, Sheet sheet) {
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			procRow(file, sheet, row);
		}
	}

	private void procRow(String file, Sheet sheet, Row row) {
		if (row == null) return;
		for (int i = 0; i < row.getLastCellNum(); i++) {
			Cell cell = row.getCell(i);
			print(file, sheet, row, cell);
		}
	}
	private void print(String file, Sheet sheet, Row row, Cell cell) {
		String value = getValue(cell);
		if (value == null) return;
		value = value.replaceAll("\\r?\\n", "\\\\n");
		System.out.println(file + ";" + sheet.getSheetName() + ";" + row.getRowNum() + ";" + cell.getColumnIndex() + ";"
				+ value);
	}
	private String getValue(Cell cell) {
		if (cell == null) return null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return "" + cell.getNumericCellValue();
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_BOOLEAN:
			return "" + cell.getBooleanCellValue();
		default:
			return null;
		}
	}

	private Workbook loadBook(File file) throws IOException, InvalidFormatException {
		FileInputStream in = new FileInputStream(file);
		try {
			Workbook book = WorkbookFactory.create(in);
			return book;
		} finally {
			if (in != null) in.close();
		}
	}
}
