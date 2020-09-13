package com.roytuts.excel.report.generation.rest.controller;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.roytuts.excel.report.generation.entity.Product;
import com.roytuts.excel.report.generation.repository.ProductRepository;

@RestController
public class ProductRestController {

	@Autowired
	private ProductRepository repository;

	@GetMapping("/report/product/")
	public ResponseEntity<Resource> generateExcelReport() throws IOException {
		List<Product> products = repository.findAll();

		Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet();

		int rowCount = 0;
		Row row = sheet.createRow(rowCount++);

		Font font = wb.createFont();
		font.setBold(true);

		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THICK);
		cellStyle.setBorderBottom(BorderStyle.THICK);
		cellStyle.setBorderLeft(BorderStyle.THICK);
		cellStyle.setBorderRight(BorderStyle.THICK);
		cellStyle.setFont(font);

		Cell cell = row.createCell(0);
		cell.setCellValue("Id");
		cell.setCellStyle(cellStyle);

		cell = row.createCell(1);
		cell.setCellValue("Name");
		cell.setCellStyle(cellStyle);

		cell = row.createCell(2);
		cell.setCellValue("Price");
		cell.setCellStyle(cellStyle);

		cell = row.createCell(3);
		cell.setCellValue("Sale Price");
		cell.setCellStyle(cellStyle);

		cell = row.createCell(4);
		cell.setCellValue("Sales Count");
		cell.setCellStyle(cellStyle);

		cell = row.createCell(5);
		cell.setCellValue("Sale Date");
		cell.setCellStyle(cellStyle);

		cellStyle = wb.createCellStyle();
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);

		for (Product product : products) {
			row = sheet.createRow(rowCount++);

			int columnCount = 0;

			cell = row.createCell(columnCount++);
			cell.setCellValue(product.getId());
			cell.setCellStyle(cellStyle);

			cell = row.createCell(columnCount++);
			cell.setCellValue(product.getName());
			cell.setCellStyle(cellStyle);

			cell = row.createCell(columnCount++);
			cell.setCellValue(product.getPrice());
			cell.setCellStyle(cellStyle);

			cell = row.createCell(columnCount++);
			cell.setCellValue(product.getSalePrice());
			cell.setCellStyle(cellStyle);

			cell = row.createCell(columnCount++);
			cell.setCellValue(product.getSalesCount());
			cell.setCellStyle(cellStyle);

			cell = row.createCell(columnCount++);
			cell.setCellValue(product.getSaleDate());
			cell.setCellStyle(cellStyle);
		}

		ByteArrayOutputStream os = new ByteArrayOutputStream();

		wb.write(os);
		wb.close();

		ByteArrayInputStream is = new ByteArrayInputStream(os.toByteArray());

		HttpHeaders headers = new HttpHeaders();
		headers.setContentType(
				MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
		headers.setCacheControl("must-revalidate, post-check=0, pre-check=0");
		headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=ProductExcelReport.xlsx");

		ResponseEntity<Resource> response = new ResponseEntity<Resource>(new InputStreamResource(is), headers,
				HttpStatus.OK);

		return response;
	}

}
