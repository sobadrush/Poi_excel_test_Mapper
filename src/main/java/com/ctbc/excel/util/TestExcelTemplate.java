package com.ctbc.excel.util;

import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;

import com.ctbc.config.RootConfig;
import com.ctbc.model.EmpDAO_Mapper;
import com.ctbc.model.EmpVO;

public class TestExcelTemplate {

	public static void main(String[] args) {
		System.setProperty("spring.profiles.active", "sqlite_env");
		
		ConfigurableApplicationContext context = new AnnotationConfigApplicationContext(RootConfig.class);

		EmpDAO_Mapper empDAO = context.getBean("empDAO_Mapper", EmpDAO_Mapper.class);

		ExcelTemplate excelTemplate = context.getBean("excelTemplate", ExcelTemplate.class);

		// 建立實體xlsx檔 (預設sheets)
		excelTemplate.doCustomerExcel(new CustomizeExceler() {
			@Override
			protected void customerExcel(Class<?> voClazz) {
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				List<String> getterMethods = ExcelUtil.getMethods(voClazz);
				List<EmpVO> empList = empDAO.getAll();// >>>>>>【查詢要輸出到Excel的資料】

				while (sheetNum < numberOfSheet) {
					int rownum = 10; // 【從第幾列開始寫】
					for (EmpVO empVO : empList) {
						Row excelRow = this.getSheets()[sheetNum].createRow(rownum++);
						int col = 0;
						for (String methodName : getterMethods) {
							Cell excelCell = excelRow.createCell(col++);
							//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
							// 字體格式
							XSSFFont font = this.getWorkbook().createFont();
							font.setColor(HSSFColor.BLACK.index); // 顏色
							font.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
							font.setFontHeightInPoints((short) 12); //字體高度
							font.setFontName("微軟正黑體"); //字體
							font.setItalic(false);   //是否使用斜體
							font.setStrikeout(false); //是否使用刪除線
							// 設定儲存格格式 
							XSSFCellStyle myCellStyle = this.getWorkbook().createCellStyle();
							myCellStyle.setFont(font); // 設定字體
							excelCell.setCellStyle(myCellStyle); // 設置單元格样式
							//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

							String getterResult = ExcelUtil.invokeGetter(empVO, methodName);// 調用VO的Getter方法
							excelCell.setCellValue(getterResult);
							this.getSheets()[sheetNum].autoSizeColumn(col - 1);// 設定每個column自動寬度
						}
					}
					//---------------
					sheetNum++;
				}

			}
		}, EmpVO.class /* 接資料的VO */, "D:/EmpData.xlsx", false /*是否覆蓋既有檔案*/);

//		// 建立實體xlsx檔 (自訂sheets)
//		excelTemplate.doCustomerExcel(new CustomizeExceler() {
//			@Override
//			protected void customerExcel(Class<?> voClazz) {
//				int sheetNum = 0; // 【從第個Sheet開始寫】
//				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
//				List<String> getterMethods = ExcelUtil.getMethods(voClazz);
//				List<CourseVO> alist = couserSvc.getCourseBy(null, null, null, "ctbc", null);// >>>>>>【查詢要輸出到Excel的資料】
//
//				while (sheetNum < numberOfSheet) {
//					int rownum = 10; // 【從第幾列開始寫】
//					for (CourseVO courseVO : alist) {
//						Row excelRow = this.getSheets()[sheetNum].createRow(rownum++);
//						int col = 0;
//						for (String methodName : getterMethods) {
//							Cell excelCell = excelRow.createCell(col++);
//							//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//							// 字體格式
//							XSSFFont font = this.getWorkbook().createFont();
//							font.setColor(HSSFColor.BLACK.index); // 顏色
//							font.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
//							font.setFontHeightInPoints((short) 12); //字體高度
//							font.setFontName("微軟正黑體"); //字體
//							font.setItalic(false);   //是否使用斜體
//							font.setStrikeout(false); //是否使用刪除線
//							// 設定儲存格格式 
//							XSSFCellStyle myCellStyle = this.getWorkbook().createCellStyle();
//							myCellStyle.setFont(font); // 設定字體
//							excelCell.setCellStyle(myCellStyle); // 設置單元格样式
//							//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
//
//							String getterResult = ExcelUtil.invokeGetter(courseVO, methodName);// 調用VO的Getter方法
//							excelCell.setCellValue(getterResult);
//							this.getSheets()[sheetNum].autoSizeColumn(col - 1);// 設定每個column自動寬度
//						}
//					}
//					//---------------
//					sheetNum++;
//				}
//
//			}
//		}, CourseVO.class /* 接資料的VO */, "D:/GGG.xlsx", new String[] { "課程總表" });

		// 取得POI workbook 的 inputStream (自訂sheets)
//		InputStream worbookInputStream = excelTemplate.doCustomerExcelGetInputStream(new CustomizeExceler() {
//			@Override
//			protected void customerExcel(Class<?> voClazz) {
//				int sheetNum = 0; // 【從第個Sheet開始寫】
//				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
//				List<String> getterMethods = ExcelUtil.getMethods(voClazz);
//				List<CourseVO> alist = couserSvc.getCourseBy(null, null, null, "ctbc", null);// >>>>>>【查詢要輸出到Excel的資料】
//
//				while (sheetNum < numberOfSheet) {
//					int rownum = 10; // 【從第幾列開始寫】
//					for (CourseVO courseVO : alist) {
//						Row excelRow = this.getSheets()[sheetNum].createRow(rownum++);
//						int col = 0;
//						for (String methodName : getterMethods) {
//							Cell excelCell = excelRow.createCell(col++);
//							//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//							// 字體格式
//							XSSFFont font = this.getWorkbook().createFont();
//							font.setColor(HSSFColor.BLACK.index); // 顏色
//							font.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
//							font.setFontHeightInPoints((short) 12); //字體高度
//							font.setFontName("微軟正黑體"); //字體
//							font.setItalic(false);   //是否使用斜體
//							font.setStrikeout(false); //是否使用刪除線
//							// 設定儲存格格式 
//							XSSFCellStyle myCellStyle = this.getWorkbook().createCellStyle();
//							myCellStyle.setFont(font); // 設定字體
//							excelCell.setCellStyle(myCellStyle); // 設置單元格样式
//							//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
//
//							String getterResult = ExcelUtil.invokeGetter(courseVO, methodName);// 調用VO的Getter方法
//							excelCell.setCellValue(getterResult);
//							this.getSheets()[sheetNum].autoSizeColumn(col - 1);// 設定每個column自動寬度
//						}
//					}
//					//---------------
//					sheetNum++;
//				}
//
//			}
//
//		}, CourseVO.class /* 接資料的VO */, new String[] { "課程總表" }/* 下方頁籤 */);

		// >>> 【輸出實體檔案測試】>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
//		int len = 2048;
//		byte[] bytes = new byte[len];
//		try (BufferedInputStream buff_is = new BufferedInputStream(worbookInputStream);
//				BufferedOutputStream buff_os = new BufferedOutputStream(new FileOutputStream(new java.io.File("D:/SSS.xlsx")))) {
//			int tmp = 0;
//			while ((tmp = buff_is.read(bytes)) != -1) {
//				System.out.println(tmp);
//				buff_os.write(bytes, 0, tmp);
//			}
//		} catch (FileNotFoundException e1) {
//			e1.printStackTrace();
//
//		} catch (IOException e1) {
//			e1.printStackTrace();
//
//		}
		// <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
		context.close();
	}

}
