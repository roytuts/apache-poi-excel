package com.roytuts.java.apache.poi.excel.deal.empty.blank.cell;

import java.util.List;

public class BlankCellHandlerTest {

	public static void main(String[] args) {
		List<Info> infoList = ExcelUtil.extractInfo("info.xlsx");

		for (Info info : infoList) {
			System.out.println(info);
		}
	}

}
