package com.roytuts.excel.poi.detect.and.delete.empty.rows;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class ExcelCellInfo {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		List<String[]> cellValues = ExcelHandler.extractInfo(new File("C:/jee_workspace/info.xlsx"));

		cellValues.forEach(c -> System.out.println(c[0] + ", " + c[1] + ", " + c[2] + ", " + c[3] + ", " + c[4]));

		ExcelHandler.writeToExcel(cellValues, new File("C:/jee_workspace/deleted_empty_rows_info.xlsx"));
	}

}
