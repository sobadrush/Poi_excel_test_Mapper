package com.ctbc.test.connection;

import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.FixMethodOrder;
import org.junit.Ignore;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.MethodSorters;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.annotation.Rollback;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.ctbc.config.RootConfig;
import com.ctbc.excel.util.CustomizeExceler;
import com.ctbc.excel.util.ExcelTemplate;
import com.ctbc.excel.util.ExcelUtil;
import com.ctbc.model.EmpDAO_Mapper;
import com.ctbc.model.EmpVO;
import com.ctbc.model.mapper.iii.EmpMapper;
/**
 * @author TizzyBac
 * 不用反射的方式產Excel
 */
@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(classes = { RootConfig.class })
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
@ActiveProfiles("sqlite_env")
public class TestExcelTemplate_HandMake {

	@Autowired
	private EmpDAO_Mapper empDAO_Mapper;

	@Autowired
	private EmpMapper empMapper;
	
	@Autowired
	private ExcelTemplate excelTemplate;

	@Test
	@Ignore
	@Rollback(true)
	public void test001() throws SQLException {
		System.out.println("============== test001 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test001 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test001 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		
		excelTemplate.doCustomerExcel(new CustomizeExceler() {
			
			private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
			
			private XSSFFont titleFont = this.getWorkbook().createFont();
			private XSSFFont contentFont = this.getWorkbook().createFont();
			private XSSFCellStyle titleCellStyle = this.getWorkbook().createCellStyle();
			private XSSFCellStyle contentCellStyle = this.getWorkbook().createCellStyle();
			private XSSFCellStyle dateCellStyle = this.getWorkbook().createCellStyle();
			
			{
				// 字體格式(TITLE列)
				titleFont.setColor(HSSFColor.BLACK.index); // 顏色
				titleFont.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
				titleFont.setFontHeightInPoints((short) 12); //字體高度
				titleFont.setFontName("標楷體"); //字體
				titleFont.setItalic(false);   //是否使用斜體
				titleFont.setStrikeout(false); //是否使用刪除線
				
				// 設定儲存格格式(TITLE列) 
				titleCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
				titleCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
				titleCellStyle.setFont(titleFont); // 設定字體
				titleCellStyle.setFillBackgroundColor(IndexedColors.CORAL.getIndex()); // 填滿顏色
				titleCellStyle.setFillPattern(CellStyle.ALIGN_FILL);// 填滿的方式
				
				// 字體格式(內容列)
				contentFont.setColor(HSSFColor.BLACK.index); // 顏色
				contentFont.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
				contentFont.setFontHeightInPoints((short) 12); //字體高度
				contentFont.setFontName("微軟正黑體"); //字體
				contentFont.setItalic(false);   //是否使用斜體
				contentFont.setStrikeout(false); //是否使用刪除線
				
				// 設定儲存格格式(內容列) 
				contentCellStyle.setFont(contentFont); // 設定字體
				contentCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
				contentCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
				contentCellStyle.setBorderBottom((short) 1);// 設定框線 
				contentCellStyle.setBorderTop((short) 1);// 設定框線 
				contentCellStyle.setBorderLeft((short) 1);// 設定框線 
				contentCellStyle.setBorderRight((short) 1);// 設定框線 
				
				// 設定[日期格式] Style
				dateCellStyle.setDataFormat((short) 14); 
				dateCellStyle.setFont(contentFont); // 設定字體
				dateCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
				dateCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
				dateCellStyle.setBorderBottom((short) 1);
				dateCellStyle.setBorderTop((short) 1);
				dateCellStyle.setBorderLeft((short) 1);
				dateCellStyle.setBorderRight((short) 1);
			}
			
			@Override
			protected void customerExcel(Class<?> voClazz) {
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				List<EmpVO> empList = empDAO_Mapper.getAll();// >>>>>>【查詢要輸出到Excel的資料】

				while (sheetNum < numberOfSheet) {
					int rownum_start = 10; // 【從第幾列開始寫】
					int colnum_start = 3;
					int colnum_title = colnum_start;
					
					// 【第0個Sheet的 [標題列] 】
					if (sheetNum == 0) {
						int titleStartRow = rownum_start - 1;
						ExcelUtil.createTitleRow(this.getSheets()[sheetNum], new String[] { "員工編號", "到職日", "姓名", "職位" }, titleCellStyle, titleStartRow, colnum_title);
					} // 標題-END

					if (sheetNum == 0) {
						// 內容-START
						for (EmpVO empVO : empList) {
							Row excelRow = this.getSheets()[sheetNum].createRow(rownum_start++);
							int colnum_content = colnum_start;
							//--------------------------------------------------------
							Cell excelCell01 = excelRow.createCell(colnum_content++);
							Cell excelCell02 = excelRow.createCell(colnum_content++);
							Cell excelCell03 = excelRow.createCell(colnum_content++);
							Cell excelCell04 = excelRow.createCell(colnum_content++);
							
							excelCell01.setCellValue(empVO.getEmpNo()); // 員工編號
//							excelCell02.setCellValue(sdf.format(empVO.getEmpHireDate()));// 到職日
							excelCell02.setCellValue(empVO.getEmpHireDate());// 到職日
							excelCell03.setCellValue(empVO.getEmpName());// 姓名
							excelCell04.setCellValue(empVO.getEmpJob());// 職位

							excelCell01.setCellStyle(contentCellStyle); // 設置單元格样式
							excelCell02.setCellStyle(dateCellStyle); // 設置單元格样式(日期格式)
							excelCell03.setCellStyle(contentCellStyle); // 設置單元格样式
							excelCell04.setCellStyle(contentCellStyle); // 設置單元格样式
							//--------------------------------------------------------
							this.getSheets()[sheetNum].autoSizeColumn(colnum_content - 1);// 設定每個column自動寬度
							this.getSheets()[sheetNum].setColumnWidth( colnum_start + 1, 5000);// 強置設定寬度
						}
					}
					//---------------
					sheetNum++;
				}

			}
		}, EmpVO.class /* 接資料的VO */, "D:/excelSample/EmpData_002.xlsx", new String[] { "員工清單001", "員工清單002" }/* excel下方的Sheet */ , true/* 是否覆蓋既有檔案 */);

	}

	
	@Test
//	@Ignore
	@Rollback(true)
	public void test002() throws SQLException {
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		
		excelTemplate.doCustomerExcel(new CustomizeExceler() {
			
			private SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
			
			private XSSFFont titleFont = this.getWorkbook().createFont();
			private XSSFFont contentFont = this.getWorkbook().createFont();
			private XSSFCellStyle titleCellStyle = this.getWorkbook().createCellStyle();
			private XSSFCellStyle contentCellStyle = this.getWorkbook().createCellStyle();
			private XSSFCellStyle dateCellStyle = this.getWorkbook().createCellStyle();
			
			{
				// 字體格式(TITLE列)
				titleFont.setColor(HSSFColor.BLACK.index); // 顏色
				titleFont.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
				titleFont.setFontHeightInPoints((short) 12); //字體高度
				titleFont.setFontName("標楷體"); //字體
				titleFont.setItalic(false);   //是否使用斜體
				titleFont.setStrikeout(false); //是否使用刪除線
				
				// 設定儲存格格式(TITLE列) 
				titleCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
				titleCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
				titleCellStyle.setFont(titleFont); // 設定字體
				titleCellStyle.setFillBackgroundColor(IndexedColors.CORAL.getIndex()); // 填滿顏色
				titleCellStyle.setFillPattern(CellStyle.ALIGN_FILL);// 填滿的方式
				
				// 字體格式(內容列)
				contentFont.setColor(HSSFColor.BLACK.index); // 顏色
				contentFont.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
				contentFont.setFontHeightInPoints((short) 12); //字體高度
				contentFont.setFontName("微軟正黑體"); //字體
				contentFont.setItalic(false);   //是否使用斜體
				contentFont.setStrikeout(false); //是否使用刪除線
				
				// 設定儲存格格式(內容列) 
				contentCellStyle.setFont(contentFont); // 設定字體
				contentCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
				contentCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
				contentCellStyle.setBorderBottom((short) 1);// 設定框線 
				contentCellStyle.setBorderTop((short) 1);// 設定框線 
				contentCellStyle.setBorderLeft((short) 1);// 設定框線 
				contentCellStyle.setBorderRight((short) 1);// 設定框線 
				
				// 設定[日期格式] Style
				dateCellStyle.setDataFormat((short) 14); 
				dateCellStyle.setFont(contentFont); // 設定字體
				dateCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
				dateCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
				dateCellStyle.setBorderBottom((short) 1);
				dateCellStyle.setBorderTop((short) 1);
				dateCellStyle.setBorderLeft((short) 1);
				dateCellStyle.setBorderRight((short) 1);
			}
			
			@Override
			protected void customerExcel(Class<?> voClazz) {
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				
				// List<EmpVO> empList = empDAO_Mapper.getAll();// >>>>>>【查詢要輸出到Excel的資料】
				List<Map<String, Object>> empList = empMapper.getEmpDeptDataMapList();// >>>>>>【查詢要輸出到Excel的資料】
				
				while (sheetNum < numberOfSheet) {
					int rownum_start = 10; // 【從第幾列開始寫】
					int colnum_start = 3;
					int colnum_title = colnum_start;
					
					// 【第0個Sheet的 [標題列] 】
					if (sheetNum == 0) {
						int titleStartRow = rownum_start - 1;
						ExcelUtil.createTitleRow(this.getSheets()[sheetNum], new String[] { "員工編號", "到職日", "姓名", "職位", "部門編號", "部門名稱" }, titleCellStyle, titleStartRow, colnum_title);
					} // 標題-END

					if (sheetNum == 0) {
						// 內容-START
						for (Map<String, Object> eMap : empList) {
							Row excelRow = this.getSheets()[sheetNum].createRow(rownum_start++);
							int colnum_content = colnum_start;
							//--------------------------------------------------------
							Cell excelCell01 = excelRow.createCell(colnum_content++);
							Cell excelCell02 = excelRow.createCell(colnum_content++);
							Cell excelCell03 = excelRow.createCell(colnum_content++);
							Cell excelCell04 = excelRow.createCell(colnum_content++);
							Cell excelCell05 = excelRow.createCell(colnum_content++);
							Cell excelCell06 = excelRow.createCell(colnum_content++);
							
							excelCell01.setCellValue(eMap.get("empno") + ""); // 員工編號
//							excelCell02.setCellValue(sdf.format(empVO.getEmpHireDate()));// 到職日
							excelCell02.setCellValue((String) eMap.get("hiredate"));// 到職日
							excelCell03.setCellValue((String) eMap.get("ename"));// 姓名
							excelCell04.setCellValue((String) eMap.get("job"));// 職位
							excelCell05.setCellValue(eMap.get("deptno") + "");// 部門編號
							excelCell06.setCellValue((String) eMap.get("dname"));// 部門名稱
							
							excelCell01.setCellStyle(contentCellStyle); // 設置單元格样式
							excelCell02.setCellStyle(dateCellStyle); // 設置單元格样式(日期格式)
							excelCell03.setCellStyle(contentCellStyle); // 設置單元格样式
							excelCell04.setCellStyle(contentCellStyle); // 設置單元格样式
							excelCell05.setCellStyle(contentCellStyle); // 設置單元格样式
							excelCell06.setCellStyle(contentCellStyle); // 設置單元格样式
							//--------------------------------------------------------
							this.getSheets()[sheetNum].autoSizeColumn(colnum_content - 1);// 設定每個column自動寬度
							this.getSheets()[sheetNum].setColumnWidth( colnum_start + 1, 5000 );// 強置設定寬度(起始寫Excel表格 +1 行，到職日的那一行)
							this.getSheets()[sheetNum].setColumnWidth( colnum_start + 3, 5000 );// 強置設定寬度(起始寫Excel表格 +3 行，職位的那一行)
						}
						
					}
					//---------------
					sheetNum++;
				}

			}
		}, null /* 接資料的VO */, "D:/excelSample/EmpData_002.xlsx", new String[] { "員工清單001", "員工清單002" }/* excel下方的Sheet */ , true/* 是否覆蓋既有檔案 */);

	}
}
