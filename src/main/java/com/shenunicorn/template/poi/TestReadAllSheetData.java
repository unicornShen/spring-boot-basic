package com.shenunicorn.template.poi;

import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestReadAllSheetData {

	public static void main(String[] args) {
		try (FileInputStream fis = new FileInputStream("C:\\Work\\TestReadAllSheetData.xlsx");) {

			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			int numberOfSheets = workbook.getNumberOfSheets();

			for (int i = 0; i < numberOfSheets; i++) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				Iterator<Row> rowsIterator = sheet.iterator();
				while (rowsIterator.hasNext()) {
					XSSFRow row = (XSSFRow) rowsIterator.next();

					Iterator<Cell> cellsIterator = row.cellIterator();
					while (cellsIterator.hasNext()) {
						XSSFCell cell = (XSSFCell) cellsIterator.next();
						System.out.println("cell = " + cell.getRawValue() + ":" + cell.toString());
					}
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
