package com.ctbc.Poi_excel_test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePOIExcelWrite {

	private static final String FILE_PATH = "/MyExcels/MyFirstExcel.xlsx";/* "/" 表示當前workspace所在的磁碟 */
	
	public static void main(String[] args) {
				
		// 建立目錄及目標檔案
		String createResult=createFolderAndFile(FILE_PATH);
		System.out.println("createResult : " + createResult);
		
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("員工資料");// Excel 下方的 sheet 的名稱
        Object[][] datatypes = {
                {"員工編號", "姓名", "分機"},
                {"z40180", "Roger1", "#3822"},
                {"z40181", "Roger2", "#3823"},
                {"z40182", "Roger3", "#3824"},
                {"z40183", "Roger4", "#3825"},
                {"z40184", "Roger5", "No fixed size"}
        };

        int rowNum = 0;
        System.out.println("Creating excel");

        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        /**************
         ** 建立檔案 **
         **************/
        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_PATH);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
        
	}//end of main
	
	
	private static String createFolderAndFile(String filePath){		
		java.io.File file = new java.io.File(filePath);
		if (file.exists() == false) {
			file.getParentFile().mkdirs(); //這樣才不會把ggg.txt當成folder建立
			try {
				file.createNewFile();
				return "檔案及資料夾建立完成!";
			} catch (IOException e) {
				return  "檔案及資料夾建立失敗!";
			}
		}else{
			return "磁碟中已有此目錄及檔案"; //已有此檔案
		}
	}
}
