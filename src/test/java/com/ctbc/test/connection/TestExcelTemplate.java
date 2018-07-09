package com.ctbc.test.connection;

import java.sql.SQLException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
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
public class TestExcelTemplate {
	
	@Autowired
	private EmpDAO_Mapper empDAO_Mapper;
	
	@Autowired
	private ExcelTemplate excelTemplate;
	
	@Test
//	@Ignore
	@Rollback(true)
	public void test001() throws SQLException {
		System.out.println("============== test001 ==============");
		
		excelTemplate.doCustomerExcel(new CustomizeExceler() {

			@Override
			protected void customerExcel(Object dao, Class<?> clazz) {
				
				EmpDAO_Mapper empDao = (EmpDAO_Mapper) dao;
				
				// TODO Auto-generated method stub
				int rownum = 999;// 【從第幾列開始寫 vs L35】
				int sheetNum = 0; // 【從第個Sheet開始寫】
				int numberOfSheet = this.getWorkbook().getNumberOfSheets();// 【共有多少個Sheet】
				List<String> getterMethods = ExcelUtil.getMethods(clazz);
				List<EmpVO> alist = empDao.getAll();

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
		}, empDAO_Mapper, EmpVO.class, "D:/參考資料/xxxGGG.xlsx", new String[] { "員工資料A", "員工資料B", "員工資料C" });
	}
}
