package com.roytuts.apachepoi.excel.cells.newlines;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCellsNewLinesApp {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		Workbook wb = new XSSFWorkbook(); // or new HSSFWorkbook();

		Sheet sheet = wb.createSheet();

		Row row = sheet.createRow(2);

		Cell cell = row.createCell(2);
		cell.setCellValue("This is first line. \n This is second line. \n This is third line.");

		// to enable newlines you need set a cell styles with wrap=true
		CellStyle cs = wb.createCellStyle();
		cs.setWrapText(true);
		cell.setCellStyle(cs);

		// increase row height to accommodate three lines of text
		row.setHeightInPoints((3 * sheet.getDefaultRowHeightInPoints()));

		// adjust column width to fit the content
		sheet.autoSizeColumn(2);

		try (OutputStream fileOut = new FileOutputStream("excel-newlines.xlsx")) {
			wb.write(fileOut);
		}

		wb.close();
	}

}
