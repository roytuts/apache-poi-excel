package com.roytuts.java.apache.poi.merge.excell.cell;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCellMerger {

	public static void main(String[] args) throws IOException {
		mergeCells();
		mergeCells("info.xlsx");
	}

	public static void mergeCells() throws IOException {
		Workbook wb = new XSSFWorkbook();

		Sheet sheet = wb.createSheet("sheet merge cells");

		Row row = sheet.createRow(1);

		Cell cell = row.createCell(1);

		cell.setCellValue("This is a test of merging cells");

		sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
				1, // last row (0-based)
				1, // first column (0-based)
				2 // last column (0-based)
		));

		// Write the output to a file
		try (OutputStream fileOut = new FileOutputStream("create-merge-cells.xlsx")) {
			wb.write(fileOut);
		}

		wb.close();
	}

	public static void mergeCells(String filename) throws IOException {
		Workbook wb = WorkbookFactory.create(new File(filename));

		Sheet sheet = wb.getSheetAt(0);

		sheet.addMergedRegion(new CellRangeAddress(2, // third row (0-based)
				3, // fourth row (0-based)
				3, // fourth column (0-based)
				4 // fifth column (0-based)
		));

		// Write the output to a file
		try (OutputStream fileOut = new FileOutputStream("existing-merge-cells.xlsx")) {
			wb.write(fileOut);
		}

		wb.close();
	}

}
