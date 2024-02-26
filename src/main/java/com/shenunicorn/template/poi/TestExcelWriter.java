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
            headerRow1.createCell(0).setCellValue("Column1");
            headerRow1.createCell(1).setCellValue("Column2");
            // 添加數據
            Row dataRow1 = sheet1.createRow(1);
            dataRow1.createCell(0).setCellValue("Data1");
            dataRow1.createCell(1).setCellValue("Data2");

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
            try (FileOutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
                workbook.write(fileOut);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
