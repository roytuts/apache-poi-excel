package com.roytuts.java.apache.poi.excel.deal.empty.blank.cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class ExcelUtil {

	private ExcelUtil() {
	}

	public static List<Info> extractInfo(final String file) {
		List<Info> infoList = new ArrayList<Info>();
		Workbook wb = null;

		try {
			wb = new XSSFWorkbook(new FileInputStream(new File(file)));
			Sheet sheet = wb.getSheetAt(0);

			boolean skipHeader = true;

			for (Row row : sheet) {
				if (skipHeader) {
					skipHeader = false;
					continue;
				}

				List<Cell> cells = new ArrayList<Cell>();

				int lastColumn = Math.max(row.getLastCellNum(), 5);// because my
																	// excel
																	// sheet has
																	// max 5
																	// columns,
																	// in case
																	// last
																	// column is
																	// empty
																	// then
																	// row.getLastCellNum()
																	// will
																	// return 4
				for (int cn = 0; cn < lastColumn; cn++) {
					Cell c = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
					cells.add(c);
				}

				Info info = extractInfoFromCell(cells);
				infoList.add(info);
			}
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				try {
					wb.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		return infoList;
	}

	private static Info extractInfoFromCell(List<Cell> cells) {
		Info info = new Info();

		Cell nameCell = cells.get(0);
		if (nameCell != null) {
			info.setName(nameCell.getStringCellValue());
		}

		Cell mobileCell = cells.get(1);
		if (mobileCell != null) {
			switch (mobileCell.getCellType()) {
			case NUMERIC:
				info.setMobile(NumberToTextConverter.toText(mobileCell.getNumericCellValue()));
				break;
			case BLANK:
				break;
			default:
				break;
			}
		}

		Cell phoneCell = cells.get(2);
		if (phoneCell != null) {
			switch (phoneCell.getCellType()) {
			case NUMERIC:
				info.setPhone(NumberToTextConverter.toText(phoneCell.getNumericCellValue()));
				break;
			case BLANK:
				break;
			default:
				break;
			}
		}

		Cell permAddressCell = cells.get(3);
		if (permAddressCell != null) {
			info.setPermAddress(permAddressCell.getStringCellValue());
		}

		Cell commAddressCell = cells.get(4);
		if (commAddressCell != null) {
			info.setCommAddress(commAddressCell.getStringCellValue());
		}

		return info;
	}

}
