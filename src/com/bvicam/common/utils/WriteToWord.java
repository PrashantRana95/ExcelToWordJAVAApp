package com.bvicam.common.utils;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.bvicam.user.dto.UserDTO;
import com.bvicam.user.helper.Helper;

public class WriteToWord {
	private static Logger logger = Logger.getLogger(WriteToWord.class);
	public static void storeInWord(String fileExcelPath, String fileWordPath) throws IOException{
		logger.debug("storeInWord Method Starts Here...");
		final int NOOFCOLS = 4;
		//ArrayList<UserDTO> userList = ExcelReader.getUsers("F:\\Prashant\\Experiments\\JAVA\\FileConverter2\\File Output\\demo.xls");
		ArrayList<UserDTO> userList = Helper.mergeSameAuthor(fileExcelPath);
//		System.out.println("Total Records are"+userList.size());
		int index = 0;
		int currentCol = 1;
		XWPFDocument document = new XWPFDocument();
		XWPFTable table = document.createTable(userList.size(),NOOFCOLS);
		table.setCellMargins(10, 200, 10, 200);
		for(UserDTO userDTO : userList){
			XWPFTableRow row  = table.getRow(index);
			if(currentCol == 1){
				XWPFParagraph para  = row.getCell(0).getParagraphs().get(0);
				para.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun r1 = para.createRun();
				r1.setFontSize(10);
				r1.setFontFamily("Calibri Light");
				
				String track = String.valueOf(userDTO.getTrack_session()).split("\\.")[0];
				r1.setText(track);
				currentCol++;
			}
			
			if(currentCol == 2){	
				XWPFParagraph para  = row.getCell(1).getParagraphs().get(0);
				para.setAlignment(ParagraphAlignment.CENTER);
				XWPFRun r1 = para.createRun();
				r1.setFontSize(10);
				r1.setFontFamily("Calibri Light");
				
				String session = String.valueOf(userDTO.getTrack_session()).split("\\.")[1];
				r1.setText(session);
				currentCol++;
			}
			
				if(currentCol == 3){	
					XWPFParagraph para  = row.getCell(2).getParagraphs().get(0);
					para.setAlignment(ParagraphAlignment.CENTER);
					XWPFRun r1 = para.createRun();
					r1.setFontSize(10);
					r1.setFontFamily("Calibri Light");
					
					r1.setText(String.valueOf(userDTO.getPaperid()));
					currentCol++;
				}
				
					if(currentCol == 4){	
						XWPFParagraph para  = row.getCell(3).getParagraphs().get(0);
						para.setAlignment(ParagraphAlignment.LEFT);
						para.setSpacingAfter(0);
						XWPFRun r1 = para.createRun();
						r1.setBold(true);
						r1.setText(String.valueOf(userDTO.getTitle_name()));
						r1.setFontSize(10);
						r1.setFontFamily("Calibri Light");
						
						XWPFParagraph paragraph = row.getCell(3).addParagraph();
			            XWPFRun r2 = paragraph.createRun();
			            paragraph.setAlignment(ParagraphAlignment.LEFT);
			            paragraph.setSpacingAfter(0);
						r2.setItalic(true);
						String sentence = String.valueOf(userDTO.getAuthor_name());
						String arr[] = sentence.split(" ");
						String authorName = "";
						for(String a : arr) {
							authorName += String.valueOf(a.charAt(0)).toUpperCase() + a.substring(1).toLowerCase()+ " ";
							authorName.trim();
						}
						r2.setText(authorName);
						r2.setFontSize(10);
						r2.setFontFamily("Calibri Light");
						
						currentCol++;
						
					}
					
					if(currentCol>4)
					{
						currentCol = 1;
						//continue;
					}
					//System.out.println(userDTO);
					index++;
			
		}
		File file = new File(fileWordPath);
		FileOutputStream fs = new FileOutputStream(file);
//		FileOutputStream fs = new FileOutputStream(new File(
//				"F:\\Prashant\\Experiments\\JAVA\\FileConverter2\\File Output\\wordfile.docx"));   
	       document.write(fs);
	       fs.close();
	       document.close();
	       System.out.println("Done...");
	       logger.debug("storeInWord Method Ends Here");
	}
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
//		storeInWord();
//		storeInWord(PropertyReader.getWordPath());
	}

}
