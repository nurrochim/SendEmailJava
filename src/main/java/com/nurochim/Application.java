package com.nurochim;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.core.io.ClassPathResource;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;

import com.nurochim.poi.excel.ExcelPOIHelper;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

@SpringBootApplication
public class Application implements CommandLineRunner {

    //https://docs.spring.io/spring/docs/5.1.6.RELEASE/spring-framework-reference/integration.html#mail
//    @Autowired
//    private JavaMailSender javaMailSender;
	

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

    void sendEmail() {

        SimpleMailMessage msg = new SimpleMailMessage();
        msg.setTo("nurochim.ristekdikti@gmail.com");

        msg.setSubject("Testing from Spring Boot");
        msg.setText("Hello World \n Spring Boot Email");

//        javaMailSender.send(msg);

    }

    void sendEmailWithAttachmentOri() throws MessagingException, IOException {

//        MimeMessage msg = javaMailSender.createMimeMessage();
//
//        // true = multipart message
//        MimeMessageHelper helper = new MimeMessageHelper(msg, true);
//        helper.setTo("nurochim.ristekdikti@gmail.com");
//
//        helper.setSubject("Testing from Spring Boot");
//
//        // default = text/plain
//        //helper.setText("Check attachment for image!");
//
//        // true = text/html
//        helper.setText("<h1>Check attachment for image!</h1>", true);
//
//        //FileSystemResource file = new FileSystemResource(new File("classpath:android.png"));
//
//        //Resource resource = new ClassPathResource("android.png");
//        //InputStream input = resource.getInputStream();
//
//        //ResourceUtils.getFile("classpath:android.png");
//
//        helper.addAttachment("06082020 Rekapitulasi WFH BBMM.docx", new File("D:\\\\Programmer Study\\\\Automated Test\\\\Email Blast\\\\06082020 Rekapitulasi WFH BBMM.docx"));

//        javaMailSender.send(msg);

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
			String emailSubject = "";
			String sendEmailStatus = "";
			
			Map<Integer, List<String>> mapRows = helperExcel.readExcel(excelPath);
			// read sheet 2 email config
			Sheet sheet2 = helperExcel.getWorkbook().getSheetAt(1);
			Row rowEmailConfig = sheet2.getRow(2);
			emailSubject = rowEmailConfig.getCell(0).getRichStringCellValue().getString()+" Test 1";
			emailMsg = rowEmailConfig.getCell(1).getRichStringCellValue().getString();
			
			for (int i = 1; i < 2; i++) {
				Row header = helperExcel.getSheet().getRow(i);
				eSertifikatCell = header.getCell(9);
				if(eSertifikatCell==null) {
					CellStyle style = helperExcel.getWorkbook().createCellStyle();
		            style.setWrapText(true);
		            
		            eSertifikatCell = helperExcel.getSheet().getRow(i).createCell(9);
		            eSertifikatCell.setCellStyle(style);
		        }
				
				List<String> rowDetails = mapRows.get(i);
				emailAddress = "";
				sertifikat = "";
				attachmentFile ="";
				sendEmailStatus = "";
				if(rowDetails.size() > 0) {
					
					emailAddress = rowDetails.get(6).trim();
					sertifikat = rowDetails.get(8).trim();
					//emailMsg = rowDetails.get(9).trim();

					attachmentFile = attachmentPath+sertifikat;
					
					SendEmailTLS sendEmailService = new SendEmailTLS();
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
