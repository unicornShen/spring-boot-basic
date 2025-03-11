package com.shenunicorn.template.poi;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestExcelWriter {
	public static void main(String[] args) {
        try (Workbook workbook = new XSSFWorkbook()) {
        	//---------------------------
            //---- 創建第一個 sheet
        	//---------------------------
            Sheet sheet1 = workbook.createSheet("Sheet1");

            // 創建表頭
            Row headerRow1 = sheet1.createRow(0);
            headerRow1.createCell(0).setCellValue("測試寬度");
            headerRow1.createCell(1).setCellValue("測試中文寬度");
            // 添加數據
            Row dataRow1 = sheet1.createRow(1);
            dataRow1.createCell(0).setCellValue("Data1");
            dataRow1.createCell(1).setCellValue("Data2");
            dataRow1.createCell(3).setCellValue("");
            
            // 以英文開度為基準 (長度 * 256)
            // 以中文寬度為基準 (長度 * 3 * 210)
            sheet1.setColumnWidth(0, 4 * 3 * 210);
            sheet1.setColumnWidth(1, 6 * 3 * 210);

            //---------------------------
            //---- 創建第二個 sheet
            //---------------------------
            Sheet sheet2 = workbook.createSheet("Sheet2");

            // 創建表頭
            Row headerRow2 = sheet2.createRow(0);
            headerRow2.createCell(0).setCellValue("ColumnA");
            headerRow2.createCell(1).setCellValue("ColumnB");
            // 添加數據
            Row dataRow2 = sheet2.createRow(1);
            dataRow2.createCell(0).setCellValue("DataA");
            dataRow2.createCell(1).setCellValue("DataB");

            // 將 workbook 寫入檔案
            try (FileOutputStream fileOut = new FileOutputStream("C:\\Work\\TestExcelWriter.xlsx")) {
                workbook.write(fileOut);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
