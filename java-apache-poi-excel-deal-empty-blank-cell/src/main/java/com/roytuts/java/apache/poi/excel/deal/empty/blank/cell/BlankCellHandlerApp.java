package com.roytuts.java.apache.poi.excel.deal.empty.blank.cell;

import java.util.List;

public class BlankCellHandlerApp {

	public static void main(String[] args) {
		List<Info> infoList = ExcelUtil.extractInfo("info.xlsx");

		infoList.forEach(i -> System.out.println(i));

		// Prior to Java 8
		// for (Info info : infoList) {
		// System.out.println(info);
		// }
	}

}
