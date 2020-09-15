package com.nurochim;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.nurochim.poi.excel.ExcelPOIHelper;

public class ReadExcelRenameAttachment {
	private static String excelPath = "D:\\Kemenristek\\Webinar Brazil\\e-sertifikat\\Penerima e-Sertifikat Webinar versi tes.xlsx";
	private static String attachmentPath = "D:\\Kemenristek\\Webinar Brazil\\e-sertifikat\\eSertifikat Peserta\\";
	
	public static void main(String[] args) {
		RenameFile renameFileHelper = new RenameFile();
		ExcelPOIHelper helperExcel = new ExcelPOIHelper();
		try {
			String name ="";
			String emailAddress ="";
			String emailMsg ="";
			String sertifikat ="";
			String oldFilePath = "";
			String newFilePath = "";
			String prefixName = "eSertifikat_";
			String ext = ".pdf";
			
			
			Map<Integer, List<String>> mapRows = helperExcel.readExcel(excelPath);
			for (int i = 1; i < mapRows.size(); i++) {
				List<String> rowDetails = mapRows.get(i);
				name="";
				sertifikat = "";
				emailAddress = "";
				emailMsg = "";
				if(rowDetails.size() > 0) {
					name = rowDetails.get(1).trim();
					sertifikat = rowDetails.get(8).trim();
					
//					emailAddress = rowDetails.get(6).trim();
//					emailMsg = rowDetails.get(9).trim();
					oldFilePath = attachmentPath+sertifikat+ext;
					newFilePath = attachmentPath+prefixName+name+ext;
					renameFileHelper.RenameFile(oldFilePath, newFilePath);
					
					// write to cell
					Row header = helperExcel.getSheet().getRow(i);
					Cell eSertifikatCell = header.getCell(8);
					eSertifikatCell.setCellValue(prefixName+name+ext);
				}
			}
			
			FileOutputStream outputStream = new FileOutputStream(excelPath);
			helperExcel.getWorkbook().write(outputStream);
			System.out.println("Rename Attachment is Successed");
		}catch (Exception e) {
			e.printStackTrace();
		}finally {
			helperExcel.closWorkbook();
		}
		
		
	}
}
