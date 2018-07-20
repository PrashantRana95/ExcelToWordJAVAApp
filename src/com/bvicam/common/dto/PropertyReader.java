package com.bvicam.common.dto;
import java.util.ResourceBundle;

public class PropertyReader {

	private static ResourceBundle rb = ResourceBundle.getBundle("config");
	
	//This is where we read the Excel Path
	public static String getExcelPath(){
		return rb.getString("excelpath");
	}
	
	//This is where we read the Word Path
	public static String getWordPath(){
		return rb.getString("wordpath");
	}
}
