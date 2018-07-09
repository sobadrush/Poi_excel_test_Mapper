package com.ctbc.excel.util;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;

import com.ctbc.config.RootConfig;
import com.ctbc.model.EmpDAO_Mapper;
import com.ctbc.model.EmpVO;

public class TestExcelTemplate {

	public static void main(String[] args) {
		ConfigurableApplicationContext context = new AnnotationConfigApplicationContext(RootConfig.class);
		EmpDAO_Mapper dao = context.getBean("empDAO_Mapper", EmpDAO_Mapper.class);

		//ExcelTemplate excelTemplate = new ExcelTemplate();
		ExcelTemplate excelTemplate = context.getBean("excelTemplate", ExcelTemplate.class);

		excelTemplate.doCustomerExcel(new CustomizeExceler() {

			@Override
			protected void customerExcel(EmpDAO_Mapper dao, Class<?> clazz) {
				// TODO Auto-generated method stub
				int rownum = 999;// 【從第幾列開始寫 vs L35】
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				List<String> getterMethods = ExcelUtil.getMethods(clazz);
				List<EmpVO> alist = dao.getAll();

				while (sheetNum < numberOfSheet) {
					rownum = 5; // 【從第幾列開始寫】
					for (EmpVO empVO : alist) {
						Row excelRow = this.getSheets()[sheetNum].createRow(rownum++);
						int col = 0;
						for (String methodName : getterMethods) {
							Cell excelCell = excelRow.createCell(col++);
							String getterResult = ExcelUtil.invokeGetter(empVO, methodName);
							excelCell.setCellValue(getterResult);
						}
					}
					//---------------
					sheetNum++;
				}
			}
		}, dao, EmpVO.class, "D:/參考資料/xxxGGG.xlsx", new String[] { "員工資料A", "員工資料B", "員工資料C" });

		context.close();
	}

}
