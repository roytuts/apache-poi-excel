package com.roytuts.excel.poi.detect.and.delete.empty.rows;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class ExcelHandler {

	private ExcelHandler() {
	}

	public static List<String[]> extractInfo(File file) {
		Workbook wb = null;

		List<String[]> infoList = new ArrayList<String[]>();

		try {
			wb = new XSSFWorkbook(new FileInputStream(file));

			Sheet sheet = wb.getSheetAt(0);

			for (Row row : sheet) {
				if (isRowEmpty(row)) {
					continue;
				}

				List<Cell> cells = new ArrayList<Cell>();

				int lastColumn = Math.max(row.getLastCellNum(), 5);

				for (int cn = 0; cn < lastColumn; cn++) {
					Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					cells.add(c);
				}

				String[] cellvalues = extractInfoFromCell(cells);
				infoList.add(cellvalues);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		return infoList;
	}

	private static String[] extractInfoFromCell(List<Cell> cells) {
		String[] cellValues = new String[5];

		cellValues[0] = getCellValue(cells.get(0));

		cellValues[1] = getCellValue(cells.get(1));

		cellValues[2] = getCellValue(cells.get(2));

		cellValues[3] = getCellValue(cells.get(3));

		cellValues[4] = getCellValue(cells.get(4));

		return cellValues;
	}

	private static String getCellValue(Cell cell) {
		String val = "";

		switch (cell.getCellType()) {
		case NUMERIC:
			val = String.valueOf(cell.getNumericCellValue());
			break;
		case STRING:
			val = cell.getStringCellValue();
			break;
		case BLANK:
			break;
		case BOOLEAN:
			val = String.valueOf(cell.getBooleanCellValue());
			break;
		case ERROR:
			break;
		case FORMULA:
			break;
		case _NONE:
			break;
		default:
			break;
		}

		return val;
	}

	private static boolean isRowEmpty(Row row) {
		boolean isEmpty = true;
		DataFormatter dataFormatter = new DataFormatter();

		if (row != null) {
			for (Cell cell : row) {
				if (dataFormatter.formatCellValue(cell).trim().length() > 0) {
					isEmpty = false;
					break;
				}
			}
		}

		return isEmpty;
	}

	public static void writeToExcel(List<String[]> cellValues, File outputFile) throws IOException {
		Workbook wb = new XSSFWorkbook();

		OutputStream outputStream = new FileOutputStream(outputFile);

		Sheet sheet = wb.createSheet();

		int rows = cellValues.size();
		int cells = cellValues.get(0).length;

		for (int i = 0; i < rows; i++) {
			Row row = sheet.createRow(i);

			for (int j = 0; j < cells; j++) {
				Cell cell = row.createCell(j);
				cell.setCellValue(cellValues.get(i)[j]);
			}
		}

		wb.write(outputStream);
		wb.close();
	}
}
