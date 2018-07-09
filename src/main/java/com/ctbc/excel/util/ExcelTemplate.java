package com.ctbc.excel.util;

import org.springframework.stereotype.Component;

@Component
public class ExcelTemplate {

	public void doCustomerExcel(CustomizeExceler exceler, Object dao, Class<?> clazz, String filePath) {
		exceler.doExcel(filePath, dao, clazz);
	}

	public void doCustomerExcel(CustomizeExceler exceler, Object dao, Class<?> clazz, String filePath, String[] sheetsnames) {
		exceler.doExcel(filePath, sheetsnames, dao, clazz);
	}
	
}
