package com.bvicam.common.utils;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.bvicam.common.dto.PropertyReader;
import com.bvicam.user.dto.UserDTO;


public class ExcelReader {

	private static Logger logger = Logger.getLogger(ExcelReader.class);
	
	public static ArrayList<UserDTO> getUsers(String filePath) throws IOException{
		logger.debug("getUsers Method Starts Here");
		ArrayList<UserDTO> userList = new ArrayList<>();
		File file = new File(filePath);
		//FirstRow is being taken false for information
		boolean firstRow = false;
		if(!file.exists()){
			throw new FileNotFoundException("File is Not Exist ");
		}
		int columnCount = 0;
		if(file.exists()){
			UserDTO userDTO = null;
			FileInputStream fs = new FileInputStream(file);
			HSSFWorkbook workBook = new HSSFWorkbook(fs);
			HSSFSheet sheet = workBook.getSheetAt(0);
			Iterator<Row> rows = sheet.rowIterator();
			while(rows.hasNext()){
				Row currentRow = rows.next();
				if(!firstRow){
					firstRow = true;
					continue;
				}
				Iterator<Cell> cells = currentRow.cellIterator();
				columnCount = 0;
				userDTO = new UserDTO();
				while(cells.hasNext()){
					Cell cell = cells.next();
					if(cell.getCellType()== Cell.CELL_TYPE_STRING){
						columnCount++;
						//Question id is set here 
						if(columnCount==1){
							userDTO.setPaperid(Integer.parseInt(cell.getStringCellValue()));
						}
						else
							if(columnCount==2){
								userDTO.setTrack_session(Double.parseDouble(cell.getStringCellValue()));
							}
							else
								if(columnCount==3){
									userDTO.setAuthor_name(cell.getStringCellValue());
								}
								else
									if(columnCount==4)
									{
										userDTO.setTitle_name(cell.getStringCellValue());
									}
					}
					else
						if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){	
							columnCount++;
							if(columnCount==1){
								
								userDTO.setPaperid((int)cell.getNumericCellValue());
							}
							else
								if(columnCount==2){
									userDTO.setTrack_session((double)cell.getNumericCellValue());
								}
						}
				}
				userList.add(userDTO);
				logger.debug("getUsers Method Ends Here...");
			}
//			System.out.println("UserList are "+userList);
			return userList;
	}
	else{
			return null;
		}
}

	public static void main(String[] args) throws IOException{
//		ArrayList<UserDTO> userList = getUsers("F:\\Prashant\\Experiments\\JAVA\\FileConverter2\\File Output\\demo.xls");
		ArrayList<UserDTO> userList = getUsers(PropertyReader.getExcelPath());
		System.out.println(userList);
		
	}
}
