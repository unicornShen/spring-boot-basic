package com.shenunicorn.template.poi;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.shenunicorn.template.poi.vo.TrunRow;

public class TestReadAllSheetData {

	public static void main(String[] args) {
		try (FileInputStream fis = new FileInputStream("C:\\Work\\TestReadAllSheetData.xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook(fis);) {

			int numberOfSheets = workbook.getNumberOfSheets();

			for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
				XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
				System.out.println("sheet.getSheetName() = " + sheet.getSheetName());

				final List<TrunRow> rowList = new ArrayList<>();

				Iterator<Row> rowsIterator = sheet.iterator();
				while (rowsIterator.hasNext()) {
					XSSFRow row = (XSSFRow) rowsIterator.next();
					int rowIndex = row.getRowNum();

					final TrunRow rowData = new TrunRow();
					if (rowIndex == 0) {
						rowData.setValueType("H");
					} else {
						rowData.setValueType("D");
					}
					
					final StringBuilder jsonValue = new StringBuilder();
					jsonValue.append("[");

					Iterator<Cell> cellsIterator = row.cellIterator();
					while (cellsIterator.hasNext()) {
						XSSFCell cell = (XSSFCell) cellsIterator.next();
						int cellIndex = cell.getColumnIndex();
						System.out.println("numberOfSheets: " + sheetIndex + " rowIndex: " + rowIndex + " cellIndex: "
								+ cellIndex + " cellValue: " + cell.toString());

						final String value = cell.toString();
						if (cellIndex == 0) {
							rowData.setAreaName(value);
						} else if (cellIndex == 1) {
							rowData.setBranchName(value);
						} else if (cellIndex == 2) {
							rowData.setEmpId(value);
						} else if (cellIndex == 2) {
							rowData.setCustId(value);
						} else {
							jsonValue.append(",");
							jsonValue.append(cell.toString());
						}
					}
					jsonValue.append("]");
					rowData.setJsonValue(StringUtils.replaceOnce(jsonValue.toString(), ",", ""));
					rowList.add(rowData);
				}
				for (TrunRow row : rowList) {
					System.out.println("numberOfSheets: " + sheetIndex + " SheetName: " + sheet.getSheetName() + " " + ToStringBuilder.reflectionToString(row, ToStringStyle.MULTI_LINE_STYLE));
					
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
