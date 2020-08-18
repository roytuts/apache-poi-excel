package com.roytuts.apache.poi.excel.color.border;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class PoiUtils {

	private PoiUtils() {
	}

	public static String getFileExtension(final String fileName) {
		if (fileName != null) {
			int len = fileName.trim().lastIndexOf(".");
			String ext = fileName.trim().substring(len);
			return ext;
		}
		return "";
	}

	public static Workbook getHSSFWorkbook() {
		return new HSSFWorkbook();
	}

	public static Workbook getXSSFWorkbook() {
		return new XSSFWorkbook();
	}

}
