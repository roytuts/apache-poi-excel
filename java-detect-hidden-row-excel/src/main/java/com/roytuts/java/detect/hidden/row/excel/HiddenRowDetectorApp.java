package com.roytuts.java.detect.hidden.row.excel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class HiddenRowDetectorApp {

	public static void main(String[] args) {
		detectHiddenRow("roytuts.xlsx");
	}

	public static void detectHiddenRow(final String fileName) {
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(new File(fileName));
			Sheet sheet = wb.getSheetAt(0);

			for (Row r : sheet) {
				if (r.getZeroHeight() || r.getHeight() == 0) {
					System.out.println("This row (" + r.getRowNum() + ") is hidden!");
				}
				for (Cell c : r) {
					System.out.print(c.getStringCellValue());
					System.out.print(" ");
				}
				System.out.println();
			}
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		} finally {
			try {
				wb.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}

}
