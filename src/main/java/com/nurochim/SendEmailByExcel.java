package com.nurochim;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import com.nurochim.poi.excel.ExcelPOIHelper;

//@SpringBootApplication
public class SendEmailByExcel implements CommandLineRunner {

	@Autowired
	private SendEmail sendEmailService;
	
    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
    }

    @Override
    public void run(String... args) {

        System.out.println("Sending Email...");

        try {
            //sendEmail();
            sendEmailWithAttachment();
        	
        } catch (Exception e) {
            e.printStackTrace();
        } 

        System.out.println("Done");
    }
    
    
	String excelPath = "D:\\Kemenristek\\Webinar Brazil\\e-sertifikat\\Penerima e-Sertifikat Webinar versi tes.xlsx";
	String attachmentPath = "D:\\Kemenristek\\Webinar Brazil\\e-sertifikat\\eSertifikat Peserta\\";
	
	public void sendEmailWithAttachment() {
		ExcelPOIHelper helperExcel = new ExcelPOIHelper();
		Cell eSertifikatCell = null;
		try {
			String emailAddress ="";
			String emailMsg ="";
			String sertifikat = "";
			String attachmentFile ="";
			String emailSubject = "Sertifikat Webinar Brazil Kemenristek/BRIN";
			String sendEmailStatus = "";
			
			Map<Integer, List<String>> mapRows = helperExcel.readExcel(excelPath);
			for (int i = 1; i < mapRows.size(); i++) {
				Row header = helperExcel.getSheet().getRow(i);
				eSertifikatCell = header.getCell(10);
				if(eSertifikatCell==null) {
					CellStyle style = helperExcel.getWorkbook().createCellStyle();
		            style.setWrapText(true);
		            
		            eSertifikatCell = helperExcel.getSheet().getRow(i).createCell(10);
		            eSertifikatCell.setCellStyle(style);
		        }
				
				List<String> rowDetails = mapRows.get(i);
				emailAddress = "";
				emailMsg = "";
				sertifikat = "";
				attachmentFile ="";
				sendEmailStatus = "";
				if(rowDetails.size() > 0) {
					
					
					emailAddress = rowDetails.get(6).trim();
					sertifikat = rowDetails.get(8).trim();
					emailMsg = rowDetails.get(9).trim();

					attachmentFile = attachmentPath+sertifikat;
							
					sendEmailStatus = sendEmailService.sendEmail(emailAddress, emailSubject, emailMsg, attachmentFile, sertifikat);
					
					
					// write to cell
					eSertifikatCell.setCellValue(sendEmailStatus);
					
				}
			}
			
			
			//System.out.println("Rename Attachment is Successed");
		}catch (Exception e) {
			e.printStackTrace();
			eSertifikatCell.setCellValue("Sending Email is Failed");
		}
		
		try {
			FileOutputStream outputStream = new FileOutputStream(excelPath);
			helperExcel.getWorkbook().write(outputStream);
		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			helperExcel.closWorkbook();
		}
		
	}	
}
