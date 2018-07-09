package com.ctbc.excel.util;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public abstract class CustomizeExceler {

	private XSSFWorkbook workbook;
	private XSSFSheet[] sheets;

	private static final int CHAR_SIZE = 256;
	private static final String DEFAULT_FILE_PATH = "/MyExcels/MyFirstExcel.xlsx"; /* "/" 表示當前workspace所在的磁碟 */

	public CustomizeExceler() {
		this.workbook = new XSSFWorkbook(); //建立 excel 檔案
		//字體格式
		XSSFFont font = workbook.createFont();
		font.setColor(HSSFColor.BLACK.index); // 顏色
		font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 粗細體

		// 設定儲存格格式 
		XSSFCellStyle styleRow1 = workbook.createCellStyle();
		styleRow1.setFillForegroundColor(HSSFColor.GREEN.index);//填滿顏色
		styleRow1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		styleRow1.setFont(font); // 設定字體
		styleRow1.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 水平置中
		styleRow1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 垂直置中

		// 設定框線 
		styleRow1.setBorderBottom((short) 1);
		styleRow1.setBorderTop((short) 1);
		styleRow1.setBorderLeft((short) 1);
		styleRow1.setBorderRight((short) 1);
		styleRow1.setWrapText(true); // 自動換行

//		/* Title */
//		sheet.autoSizeColumn(0); // 自動調整欄位寬度
//		sheet.setColumnWidth(0, CHAR_SIZE * 10);
//		sheet.setColumnWidth(1, CHAR_SIZE * 10);
//		sheet.setColumnWidth(2, CHAR_SIZE * 50);
	}

	protected void createPOI() {
		this.workbook = new XSSFWorkbook(); //建立 excel 檔案
		int defaultSheetNum = 3;
		sheets = new XSSFSheet[defaultSheetNum];// for 物件陣列
		for (int i = 0 ; i < defaultSheetNum ; i++) {
			this.sheets[i] = workbook.createSheet("工作表" + (i + 1));// 建立 Excel 下方的 sheet 
		}
	}

	protected void createPOI(String[] sheetsNames) {
		if (sheetsNames == null) {
			this.createPOI();
		} else {
			sheets = new XSSFSheet[sheetsNames.length];// for 物件陣列
			for (int i = 0 ; i < sheetsNames.length ; i++) {
				this.sheets[i] = workbook.createSheet(sheetsNames[i]);// 建立 Excel 下方的 sheet 			
			}
		}
	}

	protected void outputExcel(String file_path) {

		if ("".equals(file_path) || file_path == null ||
				file_path.length() == 0 || "dafaultPath".equals(file_path)) {
			file_path = DEFAULT_FILE_PATH;/* "/" 表示當前workspace所在的磁碟 */
		}

		// 建立目錄及目標檔案
		String createResult = ExcelUtil.createFolderAndFile(file_path);
		System.out.println("createResult : " + createResult);

		/* 將excel檔藉由 io 輸出 */
		try (FileOutputStream outputStream = new FileOutputStream(file_path);) {
			workbook.write(outputStream);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		System.out.println(" Excel 產製完成！");
	}

	protected void doExcel(String filePath, String[] sheetsNames, Object dao, Class<?> voClazz) {
		this.createPOI(sheetsNames);
		this.customerExcel(dao, voClazz);
		this.outputExcel(filePath);
	}

	protected void doExcel(String filePath, Object dao, Class<?> voClazz) {
		this.createPOI();
		this.customerExcel(dao, voClazz);
		this.outputExcel(filePath);
	}

	/* user自行實作產生Excel內容邏輯 */
	protected abstract void customerExcel(Object dao, Class<?> voClazz);

	public XSSFWorkbook getWorkbook() {
		return workbook;
	}

	public void setWorkbook(XSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public XSSFSheet[] getSheets() {
		return sheets;
	}

	public void setSheets(XSSFSheet[] sheets) {
		this.sheets = sheets;
	}

}