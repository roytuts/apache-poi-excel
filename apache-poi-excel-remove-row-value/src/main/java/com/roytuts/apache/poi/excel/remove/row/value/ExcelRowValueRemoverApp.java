package com.roytuts.apache.poi.excel.remove.row.value;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelRowValueRemoverApp {

	public static void main(String[] args) {
		final String fileName = "workbook-remove-value.xlsx";//"workbook-remove-row.xlsx";
		removeRowValueFromExcel(fileName);
	}

	public static void removeRowValueFromExcel(final String fileName) {
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

		// set column width for two columns
		sheet.setColumnWidth(0, 9000);
		sheet.setColumnWidth(1, 9000);

		// Create five rows and two columns
		for (int i = 0; i < 5; i++) {
			Row row = sheet.createRow((short) i);
			for (int j = 0; j < 2; j++) {
				row.createCell(j).setCellValue("row : " + i + ", column : " + j);
			}
		}

		// total no. of rows
		int totalRows = sheet.getLastRowNum();
		System.out.println("Total no of rows : " + totalRows);

		// remove values from third row but keep third row blank
		if (sheet.getRow(2) != null) {
			sheet.removeRow(sheet.getRow(2));
		}

		// remove third row completely - 2 for third row and +1; 2+1=3
		sheet.shiftRows(3, totalRows, -1);

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
