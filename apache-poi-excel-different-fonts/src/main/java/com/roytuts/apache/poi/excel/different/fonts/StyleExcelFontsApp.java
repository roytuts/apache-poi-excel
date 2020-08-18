package com.roytuts.apache.poi.excel.different.fonts;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class StyleExcelFontsApp {

	public static void main(String[] args) {
		final String fileName = "excel-fonts.xlsx";
		styleFontsInExcel(fileName);
	}

	public static void styleFontsInExcel(final String fileName) {
		// get the file extension
		String fileExt = PoiUtils.getFileExtension(fileName);
		Workbook workbook = null;

		// based on file extension create Workbook object
		if (".xls".equalsIgnoreCase(fileExt)) {
			workbook = PoiUtils.getHSSFWorkbook();
		} else if (".xlsx".equalsIgnoreCase(fileExt)) {
			workbook = PoiUtils.getXSSFWorkbook();
		}

		// create Sheet object
		// sheet name must not exceed 31 characters
		// the name must not contain 0x0000, 0x0003, colon(:), backslash(\),
		// asterisk(*), question mark(?), forward slash(/), opening square
		// bracket([), closing square bracket(])
		Sheet sheet = workbook.createSheet("my_sheet");

		sheet.setColumnWidth(0, 9000);
		sheet.setColumnWidth(1, 9000);
		sheet.setColumnWidth(2, 9000);

		// Create first row. Rows are 0 based.
		Row row = sheet.createRow((short) 0);

		// Create a cell
		Font font1 = workbook.createFont();
		font1.setFontHeightInPoints((short) 16);
		font1.setFontName("Courier New");
		font1.setItalic(true);
		font1.setStrikeout(true);

		CellStyle cellStyle1 = workbook.createCellStyle();
		cellStyle1.setFont(font1);

		Cell cell = row.createCell(0);
		cell.setCellValue("This is a Courier New Font");
		cell.setCellStyle(cellStyle1);

		Font font2 = workbook.createFont();
		font2.setFontHeightInPoints((short) 18);
		font2.setFontName("Arial");
		font2.setItalic(true);
		font2.setStrikeout(true);

		CellStyle cellStyle2 = workbook.createCellStyle();
		cellStyle2.setFont(font2);

		cell = row.createCell(1);
		cell.setCellValue("This is an Arial Font");
		cell.setCellStyle(cellStyle2);

		Font font3 = workbook.createFont();
		font3.setFontHeightInPoints((short) 18);
		font3.setFontName("Garamond");
		font3.setItalic(true);
		// font3.setStrikeout(true);
		CellStyle cellStyle3 = workbook.createCellStyle();
		cellStyle3.setFont(font3);

		cell = row.createCell(2);
		cell.setCellValue("This is a Garamond Font");
		cell.setCellStyle(cellStyle3);

		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(fileName);
			workbook.write(fileOut);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				fileOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

}
