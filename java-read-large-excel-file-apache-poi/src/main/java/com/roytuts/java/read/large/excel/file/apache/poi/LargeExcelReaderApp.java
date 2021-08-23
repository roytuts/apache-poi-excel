package com.roytuts.java.read.large.excel.file.apache.poi;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class LargeExcelReaderApp {

	public static void main(String[] args) throws Exception {
		String fileName = "Sales-Records.xlsx";
		// readLargeExcelFile(fileName);

		SaxEventUserModel saxEventUserModel = new SaxEventUserModel();
		saxEventUserModel.processSheets(fileName);
	}

	// The following method will give error - OutOfMemoryError
	public static void readLargeExcelFile(final String fileName) throws EncryptedDocumentException, IOException {
		Workbook wb = WorkbookFactory.create(new File(fileName));

		XSSFSheet sheet = (XSSFSheet) wb.getSheetAt(0);

		for (Row r : sheet) {
			for (Cell c : r) {
				CellType cellType = c.getCellType();
				if (CellType.STRING.equals(cellType)) {
					System.out.println(c.getStringCellValue());
				} else if (CellType.NUMERIC.equals(cellType)) {
					System.out.println(String.valueOf(c.getNumericCellValue()));
				} else if (DateUtil.isCellDateFormatted(c)) {
					System.out.println(c.getDateCellValue());
				}
			}
		}
	}
}
