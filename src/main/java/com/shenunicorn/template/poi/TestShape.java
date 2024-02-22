package com.shenunicorn.template.poi;

import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// https://www.cnblogs.com/lizhiw/p/4313789.html
public class TestShape {

	public static void main(String[] args) {
		testShape1();
	}
	
	/** OK */
	private static void testShape1() {

		try (FileOutputStream fos = new FileOutputStream("C:\\Work\\TestShape1.xls");
				HSSFWorkbook workbook = new HSSFWorkbook();) {

			HSSFSheet sheet = workbook.createSheet("Test");// 创建工作表(Sheet)
			HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
			HSSFClientAnchor anchor = new HSSFClientAnchor(255, 122, 255, 122, (short) 1, 0, (short) 4, 3);
			anchor.setCol1(1);
		    anchor.setRow1(1);
		    
			HSSFSimpleShape rec = patriarch.createSimpleShape(anchor);
			rec.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);
			rec.setLineStyle(HSSFShape.LINESTYLE_DASHGEL);// 设置边框样式
			rec.setFillColor(255, 0, 0);// 设置填充色
			rec.setLineWidth(25400);// 设置边框宽度
			rec.setLineStyleColor(0, 0, 255);// 设置边框颜色

			workbook.write(fos);

			System.out.println("end");
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println(e);
		}
	}
	
	/**
	 * 用 XSSF 目前找不到畫法...
	 */
	private static void testShape2() {

		try (FileOutputStream fos = new FileOutputStream("C:\\Work\\TestShape.xls");
				XSSFWorkbook workbook = new XSSFWorkbook();) {

			XSSFSheet sheet = workbook.createSheet("Test");// 创建工作表(Sheet)
			Drawing drawing = sheet.createDrawingPatriarch();
//			ClientAnchor anchor = new XSSFClientAnchor(255, 122, 255, 122, (short) 1, 0, (short) 4, 3);
			CreationHelper helper = workbook.getCreationHelper();
			ClientAnchor anchor = helper.createClientAnchor();
			
			// 圖片插入座標
		    anchor.setCol1(0);
		    anchor.setRow1(0);
		    

		    // 插入圖片
//		    Picture pict = drawing.createPicture(anchor, pictureIdx);

		    //圖片重新定義大小
//		    pict.resize();
			
//			XSSFSimpleShape rec = patriarch.createSimpleShape(anchor);
//			rec.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);
//			rec.setLineStyle(HSSFShape.LINESTYLE_DASHGEL);// 设置边框样式
//			rec.setFillColor(255, 0, 0);// 设置填充色
//			rec.setLineWidth(25400);// 设置边框宽度
//			rec.setLineStyleColor(0, 0, 255);// 设置边框颜色

			workbook.write(fos);

			System.out.println("end");
		} catch (Exception e) {
			e.printStackTrace();
			System.err.println(e);
		}
	}
}
