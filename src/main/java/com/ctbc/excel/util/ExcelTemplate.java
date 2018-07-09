package com.ctbc.excel.util;

import org.springframework.stereotype.Component;

import com.ctbc.model.EmpDAO_Mapper;

@Component
public class ExcelTemplate {

	public void doCustomerExcel(CustomizeExceler exceler, EmpDAO_Mapper dao, Class<?> clazz, String filePath) {
		exceler.doExcel(filePath, dao, clazz);
	}

	public void doCustomerExcel(CustomizeExceler exceler, EmpDAO_Mapper dao, Class<?> clazz, String filePath, String[] sheetsnames) {
		exceler.doExcel(filePath, sheetsnames, dao, clazz);
	}
	
}
