package com.ctbc.excel.util;

import java.io.InputStream;

import org.springframework.stereotype.Component;

/**
 * @author Z00040180
 * 
 * 【Step1】. 覆寫 getExcelInputSream()[ getExcelData() 可不寫 ] 
 * 【※※※】 調整 VO 中 private 屬性的宣告順序可改變Excel輸出時的column順序
 */
@Component
public class ExcelTemplate {

	public ExcelTemplate() {
		super();
	}

	/**
	 * 建立實體xlsx檔 (預設sheets)
	 * 
	 * @param exceler
	 * @param dao
	 * @param voClazz
	 * @param filePath
	 */
	public void doCustomerExcel(CustomizeExceler exceler, Class<?> voClazz, String filePath) {
		exceler.doExcel(filePath, voClazz);
	}

	/**
	 * 建立實體xlsx檔 (自訂sheets)
	 * 
	 * @param exceler
	 * @param dao
	 * @param voClazz
	 * @param filePath
	 * @param sheetsnames
	 */
	public void doCustomerExcel(CustomizeExceler exceler, Class<?> voClazz, String filePath, String[] sheetsnames) {
		exceler.doExcel(filePath, sheetsnames, voClazz);
	}

	/**
	 * 取得POI workbook 的 inputStream (自訂sheets)
	 * 
	 * @param exceler
	 * @param dao
	 * @param voClazz
	 * @param sheetsnames
	 */
	public InputStream doCustomerExcelGetInputStream(CustomizeExceler exceler, Class<?> voClazz, String[] sheetsnames) {
		return exceler.doExcelGetInputStream(sheetsnames, voClazz);// filePath 為傳入時，直接取得 POI workbook 的 inputStream
	}
}
