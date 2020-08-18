package com.roytuts.apache.poi.excel.color.border;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelColorBorderApp {

	public static void main(String[] args) {
		final String fileName = "excel-color-border.xlsx";
		createExcel(fileName);
	}

	public static void createExcel(final String fileName) {
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
		CellStyle cellStyle1 = workbook.createCellStyle();
		cellStyle1.setBorderTop(BorderStyle.THIN);
		cellStyle1.setTopBorderColor(IndexedColors.BLUE_GREY.getIndex());
		cellStyle1.setBorderBottom(BorderStyle.THICK);
		cellStyle1.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		cellStyle1.setBorderLeft(BorderStyle.DASHED);
		cellStyle1.setLeftBorderColor(IndexedColors.AQUA.getIndex());
		cellStyle1.setBorderRight(BorderStyle.DOTTED);
		cellStyle1.setRightBorderColor(IndexedColors.BROWN.getIndex());

		Cell cell = row.createCell(0);
		cell.setCellValue("This is surrounded with border");
		cell.setCellStyle(cellStyle1);

		// background color
		CellStyle cellStyle2 = workbook.createCellStyle();
		cellStyle2.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
		cellStyle2.setFillPattern(FillPatternType.BIG_SPOTS);
		cell = row.createCell(1);
		cell.setCellValue("Background Color");
		cell.setCellStyle(cellStyle2);

		// foreground color
		CellStyle cellStyle3 = workbook.createCellStyle();
		cellStyle3.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		cellStyle3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell = row.createCell(2);
		cell.setCellValue("Foreground Color");
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
