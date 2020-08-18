package com.roytuts.apache.poi.excel.text.alignment;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelTextAlignmentApp {

	public static void main(String[] args) {
		final String fileName = "excel-text-alignment.xlsx";
		excelTextAlignment(fileName);
	}

	public static void excelTextAlignment(final String fileName) {
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
		sheet.setColumnWidth(0, 3000);
		sheet.setColumnWidth(1, 9000);
		sheet.setColumnWidth(2, 9000);
		sheet.setColumnWidth(3, 9000);

		// Create first row. Rows are 0 based.
		Row row = sheet.createRow((short) 0);

		// Create a cell
		// put a value in cell.
		CellStyle cellStyle1 = workbook.createCellStyle();
		cellStyle1.setWrapText(true);

		// justify text alignment
		cellStyle1.setAlignment(HorizontalAlignment.JUSTIFY);
		Cell cell = row.createCell(0);
		cell.setCellValue("This is Justify Alignment");
		cell.setCellStyle(cellStyle1);

		CellStyle cellStyle2 = workbook.createCellStyle();
		cellStyle2.setWrapText(true);

		// text left alignment
		cellStyle2.setAlignment(HorizontalAlignment.LEFT);
		cell = row.createCell(1);
		cell.setCellValue("This is Left Alignment");
		cell.setCellStyle(cellStyle2);

		CellStyle cellStyle3 = workbook.createCellStyle();
		cellStyle3.setWrapText(true);

		// text right alignment
		cellStyle3.setAlignment(HorizontalAlignment.RIGHT);
		cell = row.createCell(2);
		cell.setCellValue("This is Right Alignment");
		cell.setCellStyle(cellStyle3);

		CellStyle cellStyle4 = workbook.createCellStyle();
		cellStyle4.setWrapText(true);

		// text center alignment
		cellStyle4.setAlignment(HorizontalAlignment.CENTER);
		cell = row.createCell(3);
		cell.setCellValue("This is Center Alignment");
		cell.setCellStyle(cellStyle4);

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
