package com.ctbc.excel.util;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class ExcelUtil {

	public static String createFolderAndFile(String filePath) {
		java.io.File file = new java.io.File(filePath);
		if (file.exists() == false) {
			file.getParentFile().mkdirs(); //這樣才不會把ggg.txt當成folder建立
			try {
				file.createNewFile();
				return "檔案及資料夾建立完成! ( " + filePath + " )";
			} catch (IOException e) {
				return "檔案及資料夾建立失敗! ( " + filePath + " )";
			}
		} else {
			return "磁碟中已有此目錄及檔案 ( " + filePath + " )"; //已有此檔案
		}
	}

	/**
	 * @param clazz  : vo.class
	 * @param String : methodStartWith
	 * @return List<類別中方法的名稱字串> List內容的順序根據VO中宣告private屬性的順序排序
	 */
	public static List<String> getMethods(Class<?> clazz, String methodStartWith) {
		List<String> methodList = new ArrayList<String>();
		Field[] fields = clazz.getDeclaredFields();   // 取得屬性
		Method[] methods = clazz.getDeclaredMethods();// 取得所有方法

		for (Field field : fields) {
			for (Method mm : methods) {
				String methodName = mm.getName();
				if (methodName.startsWith(methodStartWith)) {// 取得 methodStartWith字串開頭 的方法的名稱 : 加到 methodList
					String methodNameWithoutGet /* 對方法名稱作字串切割 =>取得 getter方法名稱不含 "get" 的字串 */
						= methodName.substring(methodName.indexOf(methodStartWith) + (methodStartWith.length()));
					if (field.getName().equalsIgnoreCase(methodNameWithoutGet)) {
						methodList.add(methodName);
					}
				}
			}
		}
		return methodList;
	}

	/**
	 * @param clazz  : vo.class
	 * @param String : methodStartWith，預設值 : "get"
	 * @return List<類別中方法的名稱字串> List內容的順序根據VO中宣告private屬性的順序排序
	 */
	public static List<String> getMethods(Class<?> clazz /*, String methodStartWith */) {
		return getMethods(clazz,"get");
	}
	
	/**
	 * @param valueObject : vo 物件   
	 * @param methodName : 方法名稱字串       
	 * @return 呼叫getter的結果字串
	 */
	public static String invokeGetter(Object valueObject, String methodName) {
		String getterResultString = null;
		try {
			getterResultString = valueObject.getClass().getMethod(methodName).invoke(valueObject).toString();
		} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException | NoSuchMethodException | SecurityException e) {
			e.printStackTrace();
		}
		return getterResultString;
	}
	
	public static void main(String[] args) {
//		Field[] fields = EmpVO.class.getDeclaredFields();
//		for (Field field : fields) {
//			System.out.println(field.toString());
//		}

//		List<String> fmethods = ExcelUtil.getMethods(EmpVO.class, "get");
//		for (String str : fmethods) {
//			System.out.println(str.toString());
//		}
	}

}
