package com.ctbc.Poi_excel_test;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;

public class FileWriterTest {

	public static void main(String[] args) {

		FileWriter fWriter = null;
		File file = new File("D:/testFolder/ttt2/ggg.txt");
		if (file.exists() == false) {
			file.getParentFile().mkdir();
			try {
				file.createNewFile();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		
		try {
			fWriter = new FileWriter(new File("D:/testFolder/ggg.txt"));
			fWriter.write("fuckqweqewe");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				fWriter.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}

}
