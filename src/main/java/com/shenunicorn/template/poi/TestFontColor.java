package com.shenunicorn.template.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestFontColor {

	public static void main(String[] args) {
		System.out.println("--- Start ---");
		try (FileOutputStream fos = new FileOutputStream("C:\\Work\\TestColor1.xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook();) {

			testColor1(workbook);

			workbook.write(fos);
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println(e);
		}

		System.out.println("--- End ---");
	}

	/** OK */
	private static void testColor1(XSSFWorkbook workBook) {

		XSSFSheet sheet = workBook.createSheet("Test Color");

		XSSFFont font = workBook.createFont();
		font.setColor(HSSFColorPredefined.RED.getIndex());
		font.setFontHeight(20); // 字型大小 (點下去才生效?)

		Row row = sheet.createRow(0);
		row.setHeight((short) (2 * 512));

		CellStyle cellStyle = workBook.createCellStyle();
		cellStyle.setWrapText(true); // 自動換行

		Cell cell = row.createCell(0);
		cell.setCellStyle(cellStyle);

		XSSFRichTextString richText = new XSSFRichTextString("● \n (111)");
		richText.applyFont(0, 1, getFont(workBook, HSSFColorPredefined.RED));
		cell.setCellValue(richText);

		//
		Cell cell2 = row.createCell(1);
		cell2.setCellStyle(cellStyle);
		XSSFRichTextString richText2 = new XSSFRichTextString("● \n (222)");
		richText2.applyFont(0, 1, getFont(workBook, HSSFColorPredefined.YELLOW));
		cell2.setCellValue(richText2);

		//
		Cell cell3 = row.createCell(2);
		cell3.setCellStyle(cellStyle);
		XSSFRichTextString richText3 = new XSSFRichTextString("● \n (333)");
		richText3.applyFont(0, 1, getFont(workBook, HSSFColorPredefined.GREEN));
		cell3.setCellValue(richText3);
	}

	private static XSSFFont getFont(XSSFWorkbook workbook, HSSFColorPredefined color) {
		XSSFFont font = workbook.createFont();
		font.setColor(color.getIndex());
//		font.setFontHeight(36); // 字型大小 (點下去才生效?)
		font.setFontHeightInPoints((short) 20); // 字型大小 (點下去才生效?)
		font.setBold(true);

		return font;
	}
}
