package com.shenunicorn.template.poi;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestMergedRegion {

	public static void main(String[] args) {
		try (FileOutputStream fos = new FileOutputStream("C:\\Work\\TestMergedRegion.xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook();) {

			Sheet sheet = workbook.createSheet();
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 4, 5));

			Row row = sheet.createRow(0);
			Cell cell_1 = row.createCell(3);

			cell_1.setCellValue("[1] test value");

			// cell 位置3-9被合併成一個儲存格，不管你怎麼建立第4個cell還是第5個cell…然後在寫資料。 都是無法寫入的。
			Cell cell_2 = row.createCell(10);

			cell_2.setCellValue("[2] test value");
			
			

			workbook.write(fos);
			
			System.out.println("end");
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println(e);
		}
	}
}
