package com.ctbc.test.connection;

import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.FixMethodOrder;
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

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(classes = { RootConfig.class })
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
@ActiveProfiles("sqlite_env")
public class TestExcelTemplate_HandMake {

	@Autowired
	private EmpDAO_Mapper empDAO_Mapper;

	@Autowired
	private ExcelTemplate excelTemplate;

	@Test
//	@Ignore
	@Rollback(true)
	public void test002() throws SQLException {
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		
		excelTemplate.doCustomerExcel(new CustomizeExceler() {
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
						//>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
						// 字體格式
						XSSFFont font = this.getWorkbook().createFont();
						font.setColor(HSSFColor.BLACK.index); // 顏色
						font.setBoldweight(Font.BOLDWEIGHT_NORMAL); // 粗細體
						font.setFontHeightInPoints((short) 12); //字體高度
						font.setFontName("標楷體"); //字體
						font.setItalic(false);   //是否使用斜體
						font.setStrikeout(false); //是否使用刪除線
						// 設定儲存格格式 
						XSSFCellStyle myCellStyle = this.getWorkbook().createCellStyle();
						myCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
						myCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
						myCellStyle.setFont(font); // 設定字體

						myCellStyle.setFillBackgroundColor(IndexedColors.CORAL.getIndex()); // 填滿顏色
						myCellStyle.setFillPattern(CellStyle.ALIGN_FILL);// 填滿的方式
						//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
						ExcelUtil.createTitleRow(this.getSheets()[sheetNum], new String[] { "員工編號", "到職日", "姓名", "職位" }, myCellStyle, titleStartRow, colnum_title);
					} // 標題-END

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
					myCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
					myCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
					// 設定框線 
					myCellStyle.setBorderBottom((short) 1);
					myCellStyle.setBorderTop((short) 1);
					myCellStyle.setBorderLeft((short) 1);
					myCellStyle.setBorderRight((short) 1);
					//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
					
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

							// 設定[日期格式]Style
							XSSFCellStyle myCellStyleDate = this.getWorkbook().createCellStyle();
							myCellStyleDate.setDataFormat((short) 14); 
							myCellStyleDate.setFont(font); // 設定字體
							myCellStyleDate.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
							myCellStyleDate.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
							myCellStyleDate.setBorderBottom((short) 1);
							myCellStyleDate.setBorderTop((short) 1);
							myCellStyleDate.setBorderLeft((short) 1);
							myCellStyleDate.setBorderRight((short) 1);
							
							excelCell01.setCellStyle(myCellStyle); // 設置單元格样式
							excelCell02.setCellStyle(myCellStyleDate); // 設置單元格样式(日期格式)
							excelCell03.setCellStyle(myCellStyle); // 設置單元格样式
							excelCell04.setCellStyle(myCellStyle); // 設置單元格样式
							//--------------------------------------------------------
							this.getSheets()[sheetNum].autoSizeColumn(colnum_content - 1);// 設定每個column自動寬度
							this.getSheets()[sheetNum].setColumnWidth( colnum_start + 1, 5000);
						}
					}
					//---------------
					sheetNum++;
				}

			}
		}, EmpVO.class /* 接資料的VO */, "D:/excelSample/EmpData_002.xlsx", new String[] { "員工清單001", "員工清單002" }/* excel下方的Sheet */ , true/* 是否覆蓋既有檔案 */);

	}

}
