package com.bvicam.user.helper;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Properties;

import org.apache.log4j.Logger;

import com.bvicam.common.dto.PropertyReader;
import com.bvicam.common.utils.ExcelReader;
import com.bvicam.user.dto.UserDTO;

public class Helper {
	private static Logger logger = Logger.getLogger(Helper.class);
	
	public static ArrayList<UserDTO> fitAnd(ArrayList<UserDTO> userList){
		for(UserDTO userDTO : userList){
			String authorName = userDTO.getAuthor_name();
			final String REPLACEWITH = "and";
			int index = authorName.lastIndexOf(",");
			if(index>=0){
				StringBuilder sb = new StringBuilder(authorName);
				sb.deleteCharAt(index);
				sb.insert(index, REPLACEWITH);
//				sb.insert(index, "and");
				userDTO.setAuthor_name(sb.toString());
			}		
//		System.out.println("Author "+authorName+" ::: "+userDTO.getAuthor_name());
		}
		return userList;
	}
	
	public static ArrayList<UserDTO> mergeSameAuthor(String filePath) throws IOException{
		logger.debug("mergeSameAuthor Method Starts Here");
		File file = new File(filePath);
		FileInputStream fs = new FileInputStream(file);
		ArrayList<UserDTO> userList = ExcelReader.getUsers(filePath);
//		ArrayList<UserDTO> userList = ExcelReader.getUsers("F:\\Prashant\\Experiments\\JAVA\\FileConverter2\\File Output\\demo.xls");
//		System.out.println("Total Records are "+userList.size());
		LinkedHashMap<Integer,UserDTO> map = new LinkedHashMap<>();
		for(UserDTO userDTO : userList){
			if(map.containsKey(userDTO.getPaperid())){
				UserDTO userDTOObject = map.get(userDTO.getPaperid());
				userDTOObject.setAuthor_name(userDTOObject.getAuthor_name() + " , "+ userDTO.getAuthor_name());
			}
			else{
				map.put(userDTO.getPaperid(), userDTO);
			}
		}
		System.out.println(userList.size());
		System.out.println(map.size());
		ArrayList<UserDTO> updateList  = new ArrayList<>();
		for(Map.Entry<Integer, UserDTO> m: map.entrySet()){
			updateList.add(m.getValue());
		}
		logger.debug("mergeSameAuthor Method Ends Here....");
		updateList = fitAnd(updateList);
		return updateList;
	}

	public static void main(String args[]) throws IOException
	{
//		System.out.println("After "+mergeSameAuthor());
		System.out.println("After "+mergeSameAuthor(PropertyReader.getWordPath()));
	}
	
}
