package com.nurochim;

import java.io.File;
import java.util.Properties;

import javax.annotation.PostConstruct;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;

@Service
public class SendEmail {
	
	 @Autowired
	 private JavaMailSender javaMailSender;
	 
	 @Autowired
	 public SendEmail(JavaMailSender mailSender) {
	     this.javaMailSender = mailSender;
	 }
 
	public String sendEmail(String emailAddress, String emailSubject, String emailMsg, String attachmentFile, String attachmentName) {
		try {
			// send email with attachment
			MimeMessage msg = javaMailSender.createMimeMessage();
			msg.getSession();
	        // true = multipart message
	        MimeMessageHelper helper = new MimeMessageHelper(msg, true);
	        helper.setTo(emailAddress);
	        helper.setSubject(emailSubject);
	        helper.setText(emailMsg, true);
	        helper.addAttachment(attachmentName, new File(attachmentFile));

	        javaMailSender.send(msg);
	        return "Sending Email is Succes";
		} catch (Exception e) {
			e.printStackTrace();
			return "Sending Email is Failed";
		}
	}
	
	
	public String sendEmailMimeMsg() {
		// Recipient's email ID needs to be mentioned.
	      String to = "abcd@gmail.com";

	      // Sender's email ID needs to be mentioned
	      String from = "web@gmail.com";

	      // Assuming you are sending email from localhost
	      String host = "localhost";

	      // Get system properties
	      Properties properties = System.getProperties();

	      // Setup mail server
	      properties.setProperty("mail.smtp.host", host);

	      // Get the default Session object.
	      Session session = Session.getDefaultInstance(properties);

	      try {
	         // Create a default MimeMessage object.
	         MimeMessage message = new MimeMessage(session);

	         // Set From: header field of the header.
	         message.setFrom(new InternetAddress(from));

	         // Set To: header field of the header.
	         message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));

	         // Set Subject: header field
	         message.setSubject("This is the Subject Line!");

	         // Now set the actual message
	         message.setText("This is actual message");

	         // Send message
	         Transport.send(message);
	         System.out.println("Sent message successfully....");
	         return "Sending Email is Succes";
	      } catch (MessagingException mex) {
	         mex.printStackTrace();
	         return "Sending Email is Failed";
	      }
	}
}
