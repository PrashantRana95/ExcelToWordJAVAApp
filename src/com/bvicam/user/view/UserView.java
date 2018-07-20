package com.bvicam.user.view;
import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.sql.rowset.serial.SerialException;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.log4j.Logger;

import com.bvicam.common.utils.ExcelReader;
import com.bvicam.common.utils.WriteToWord;
import com.bvicam.user.helper.Helper;

public class UserView extends JFrame {

	private JPanel contentPane;
	private JTextField textField;
	private JTextField textField_1;
	private Logger logger = Logger.getLogger(UserView.class);
	JButton btnExcel = new JButton("Add Excel File");
	JButton btnWordConverter = new JButton("Conert To Word");
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UserView frame = new UserView();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public UserView() {
		logger.debug("UserView Constructor Call Start...");
		setForeground(Color.WHITE);
		setLocationRelativeTo(null);
		setResizable(false);
		
		setTitle("ExcelToWordConverter");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 567, 702);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		setLocationRelativeTo(null);
		
		contentPane.setLayout(null);
		
		textField = new JTextField();
		textField.setBounds(116, 208, 373, 31);
		contentPane.add(textField);
		textField.setColumns(10);
		
		
		btnExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					newFileOpen();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		btnExcel.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/AddFile.png")));
		btnExcel.setBounds(203, 262, 178, 52);
		contentPane.add(btnExcel);
		
		
		btnWordConverter.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					openWordFile();
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		btnWordConverter.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/ConvertFile.png")));
		btnWordConverter.setBounds(203, 490, 180, 52);
		contentPane.add(btnWordConverter);
		
		textField_1 = new JTextField();
		textField_1.setBounds(116, 434, 373, 31);
		contentPane.add(textField_1);
		textField_1.setColumns(10);
		
		JLabel label = new JLabel("");
		label.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/FileConverterLogo.png")));
		label.setBounds(33, 0, 124, 95);
		contentPane.add(label);
		
		JLabel lblStep = new JLabel("Step 1 : Select Excel File");
		lblStep.setFont(new Font("Bernard MT Condensed", Font.PLAIN, 17));
		lblStep.setBounds(222, 134, 220, 31);
		contentPane.add(lblStep);
		
		JLabel label_1 = new JLabel("");
		label_1.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/xlsFile.png")));
		label_1.setBounds(156, 120, 56, 62);
		contentPane.add(label_1);
		
		JLabel label_2 = new JLabel("");
		label_2.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/docxFile.png")));
		label_2.setBounds(156, 336, 56, 70);
		contentPane.add(label_2);
		
		JLabel lblStep_1 = new JLabel("Step 2 : Convert The Excel To Word File");
		lblStep_1.setFont(new Font("Bernard MT Condensed", Font.PLAIN, 17));
		lblStep_1.setBounds(222, 363, 299, 31);
		contentPane.add(lblStep_1);
		
		JLabel label_3 = new JLabel("");
		label_3.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/MainFileLogo.png")));
		label_3.setBounds(105, 11, 367, 74);
		contentPane.add(label_3);
		
		JLabel label_4 = new JLabel("");
		label_4.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/FilePath.png")));
		label_4.setBounds(50, 191, 56, 62);
		contentPane.add(label_4);
		
		JLabel label_5 = new JLabel("");
		label_5.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/DocFilePath.png")));
		label_5.setBounds(50, 419, 56, 62);
		contentPane.add(label_5);
		
		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/brainmentors.png")));
		lblNewLabel.setBounds(0, 555, 268, 110);
		contentPane.add(lblNewLabel);
		
		JLabel label_6 = new JLabel("");
		label_6.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/gpr.png")));
		label_6.setBounds(434, 576, 117, 95);
		contentPane.add(label_6);
		
		JLabel label_7 = new JLabel("");
		label_7.setIcon(new ImageIcon(UserView.class.getResource("/com/bvicam/images/logoback.png")));
		label_7.setBounds(376, 576, 100, 100);
		contentPane.add(label_7);
		logger.debug("UserView Constructor Call Ends Here");
	}
	
	
	
	public static String fileDirectory = "";
	private void newFileOpen() throws IOException{	
		logger.debug("newFileOpen Method Starts");
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("Please Open Excel File Only");
//		 fileChooser.addChoosableFileFilter(new FileFilter() {
//			@Override
//			public String getDescription() {
//				// TODO Auto-generated method stub
//				return "Excel Documents Only(*.xls)";
//			}
//			
//			@Override
//			public boolean accept(File f) {
//				// TODO Auto-generated method stub
//				if(f.isDirectory()){
//					return true;
//				}
//				else{
//					return f.getName().toLowerCase().endsWith(".xls");
//				}
//			}
//		});
//		fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		FileNameExtensionFilter xlsFilter = new FileNameExtensionFilter("Excel Documents Only(*.xls)", "xls");
//		fileChooser.addChoosableFileFilter(xlsFilter);
		fileChooser.setFileFilter(xlsFilter);
		textField.setEditable(false);
        int returnValue = fileChooser.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile().getAbsoluteFile();
            fileDirectory = selectedFile.getAbsolutePath();
            if(fileDirectory.endsWith("xls")){
                textField.setText(fileChooser.getSelectedFile().getAbsolutePath()); 
//                ExcelReader excelReader = new ExcelReader();
                ExcelReader.getUsers(fileDirectory);
//                Helper helper = new Helper();
                Helper.mergeSameAuthor(fileDirectory);
                JOptionPane.showMessageDialog(this, "Excel File Uploaded SuccessFully");
//                System.out.println("Helper Path"+helper.mergeSameAuthor(fileDirectory));
//                System.out.println("Excel Path "+excelReader.getUsers(fileDirectory));
            }
            else{
            	JOptionPane.showMessageDialog(this, "Please Select Excel File With .xls Extension Only");
            }
        }
        else if (returnValue == JFileChooser.CANCEL_OPTION){
        	JOptionPane.showMessageDialog(this , "File Not Selected");
        }
        logger.debug("newFileOpen Method Ends Here....");
	}
	
	private void openWordFile() throws IOException{
		logger.debug("openWordFile Method Starts...");
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.setDialogTitle("Specify The Filename To be Saved");
		FileNameExtensionFilter docFilter = new FileNameExtensionFilter("Word Documents Only(*.docx)", "docx");
		fileChooser.setFileFilter(docFilter);
//		fileChooser.addChoosableFileFilter(new FileFilter() {
//			
//			@Override
//			public String getDescription() {
//				// TODO Auto-generated method stub
//				return "Word Documents Only(*.docx)";
//			}
//			
//			@Override
//			public boolean accept(File f) {
//				// TODO Auto-generated method stub
//				if(f.isDirectory()){
//					return true;
//				}
//				else{
//					return f.getName().toLowerCase().endsWith(".docx");
//				}
//			}
//		});;
		int returnValue = fileChooser.showSaveDialog(null);
		if(returnValue == JFileChooser.APPROVE_OPTION){
			textField_1.setEditable(false);
			String filePath = textField.getText();
			//If Excel File Not Selected then this Message Will Pop Out
			if(!(filePath.length()>0)){
				JOptionPane.showMessageDialog(this, "Please Select Excel File First to Begin the Conversion Process");
			}
			else if (filePath.length()>0){
				String fileExtension = ".docx";
				System.out.println(filePath);
//				fileChooser.showSaveDialog(this);
				File selectedFile = fileChooser.getSelectedFile().getAbsoluteFile();
//				String wordFilePath = selectedFile.getAbsolutePath().concat(fileExtension);
//				WriteToWord writeWord = new WriteToWord();
				if(!(selectedFile.getAbsolutePath().endsWith(fileExtension))){
					String wordFilePath = selectedFile.getAbsolutePath().concat(fileExtension);
					WriteToWord.storeInWord(filePath , wordFilePath);
					textField_1.setText(wordFilePath);	
				}
				else if(selectedFile.getAbsolutePath().endsWith(fileExtension))
				{
					String wordFilePath = selectedFile.getAbsolutePath();
					WriteToWord.storeInWord(filePath, wordFilePath);
					textField_1.setText(wordFilePath);
				}
				JOptionPane.showMessageDialog(this, "Excel To Word File Converted Successfully");
			}
		}
		else if (returnValue == JFileChooser.CANCEL_OPTION){
			JOptionPane.showMessageDialog(this, "No Options Selected...File Conversion Not Completed");
		}
		logger.debug("openWordFile Method Ends Here");
	}
}
