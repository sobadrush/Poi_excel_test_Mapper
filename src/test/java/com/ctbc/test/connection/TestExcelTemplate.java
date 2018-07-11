package com.ctbc.test.connection;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
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

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(classes = { RootConfig.class })
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
@ActiveProfiles("sqlite_env")
public class TestExcelTemplate {

	@Autowired
	private EmpDAO_Mapper empDAO_Mapper;

	@Autowired
	private ExcelTemplate excelTemplate;

	@Test
	@Ignore
	@Rollback(true)
	public void test001() throws SQLException {
		System.out.println("============== test001 >>> 建立實體xlsx檔 (預設sheets) ==============");
		System.out.println("============== test001 >>> 建立實體xlsx檔 (預設sheets) ==============");
		System.out.println("============== test001 >>> 建立實體xlsx檔 (預設sheets) ==============");

		// 建立實體xlsx檔 (預設sheets)
		excelTemplate.doCustomerExcel(new CustomizeExceler() {
			@Override
			protected void customerExcel(Class<?> voClazz) {
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				List<String> getterMethods = ExcelUtil.getMethods(voClazz);
				List<EmpVO> empList = empDAO_Mapper.getAll();// >>>>>>【查詢要輸出到Excel的資料】

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

	}

	@Test
//	@Ignore
	@Rollback(true)
	public void test002() throws SQLException {
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");
		System.out.println("============== test002 >>> 建立實體xlsx檔 (自訂sheets) ==============");

		excelTemplate.doCustomerExcel(new CustomizeExceler() {
			@Override
			protected void customerExcel(Class<?> voClazz) {
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				List<String> getterMethods = ExcelUtil.getMethods(voClazz);
				List<EmpVO> empList = empDAO_Mapper.getAll();// >>>>>>【查詢要輸出到Excel的資料】

				while (sheetNum < numberOfSheet) {
					int rownum = 10; // 【從第幾列開始寫】
					int colnum_start = 3;
					
					// 【第0個Sheet的 [標題列] 】
					if (sheetNum == 0) {
						int titleStartRow = rownum - 1;
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
						ExcelUtil.createTitleRow(this.getSheets()[sheetNum], new String[] { "員工編號", "到職日", "姓名", "職位" }, myCellStyle, titleStartRow, colnum_start);
					} // 標題-END

					// 內容-START
					for (EmpVO empVO : empList) {
						Row excelRow = this.getSheets()[sheetNum].createRow(rownum++);
						int col = colnum_start;
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
							myCellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
							myCellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中
							excelCell.setCellStyle(myCellStyle); // 設置單元格样式
							//<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

							// 【直接反射VO中的getter】
							String getterResult = ExcelUtil.invokeGetter(empVO, methodName);// 調用VO的Getter方法
							excelCell.setCellValue(getterResult);
							
							this.getSheets()[sheetNum].autoSizeColumn(col - 1);// 設定每個column自動寬度
						}
					}
					//---------------
					sheetNum++;
				}

			}
		}, EmpVO.class /* 接資料的VO */, "D:/excelSample/EmpData_002.xlsx", new String[] { "員工清單001", "員工清單002" }/* excel下方的Sheet */ , true/*是否覆蓋既有檔案*/);

	}

	@Test
	@Ignore
	@Rollback(true)
	public void test003() throws SQLException {
		System.out.println("============== test003 >>> 取得POI workbook 的 inputStream (自訂sheets) ==============");
		System.out.println("============== test003 >>> 取得POI workbook 的 inputStream (自訂sheets) ==============");
		System.out.println("============== test003 >>> 取得POI workbook 的 inputStream (自訂sheets) ==============");

		InputStream worbookInputStream = excelTemplate.doCustomerExcelGetInputStream(new CustomizeExceler() {
			@Override
			protected void customerExcel(Class<?> voClazz) {
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				List<String> getterMethods = ExcelUtil.getMethods(voClazz);
				List<EmpVO> empList = empDAO_Mapper.getAll();// >>>>>>【查詢要輸出到Excel的資料】

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

		}, EmpVO.class /* 接資料的VO */, new String[] { "員工清單001" }/* excel下方的Sheet */);

		// >>> 【輸出實體檔案測試】>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
		int len = 2048;
		byte[] bytes = new byte[len];
		try (BufferedInputStream buff_is = new BufferedInputStream(worbookInputStream);
				BufferedOutputStream buff_os = new BufferedOutputStream(new FileOutputStream(new java.io.File("D:/EmpData_003.xlsx")))) {
			int readedLen = 0;
			while ((readedLen = buff_is.read(bytes)) != -1) {
				System.out.println(readedLen + " kb");
				buff_os.write(bytes, 0, readedLen);
			}
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		// <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	}

}
