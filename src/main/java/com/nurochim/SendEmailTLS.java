package com.nurochim;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import java.util.Properties;

public class SendEmailTLS {
	public String sendEmail(String emailAddress, String emailSubject, String emailMsg, String attachmentFile, String attachmentName) {

        final String username = "nurochim.ristekbrin@gmail.com";
        final String password = "vnoibmlaqyzocegs";

        Properties prop = new Properties();
		prop.put("mail.smtp.host", "smtp.gmail.com");
        prop.put("mail.smtp.port", "587");
        prop.put("mail.smtp.auth", "true");
        prop.put("mail.smtp.starttls.enable", "true"); //TLS

        Session session = Session.getInstance(prop,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });

        try {

            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress("nurochim.ristekbrin@gmail.com"));
            message.setRecipients(
                    Message.RecipientType.TO,
                    InternetAddress.parse(emailAddress)
            );
            message.setSubject(emailSubject);
//            message.setText(emailMsg);
            
            // attachment
            Multipart multipart = new MimeMultipart();

            MimeBodyPart textBodyPart = new MimeBodyPart();
//            textBodyPart.setText(emailMsg);
            textBodyPart.setContent(emailMsg, "text/html; charset=UTF-8");

            MimeBodyPart attachmentBodyPart= new MimeBodyPart();
            DataSource source = new FileDataSource(attachmentFile); // ex : "C:\\test.pdf"
            attachmentBodyPart.setDataHandler(new DataHandler(source));
            attachmentBodyPart.setFileName(attachmentName); // ex : "test.pdf"

            multipart.addBodyPart(textBodyPart);  // add the text part
            multipart.addBodyPart(attachmentBodyPart); // add the attachement part

            message.setContent(multipart);
            

            Transport.send(message);

            System.out.println("Send email to "+emailAddress+" is Done");
            return "Sending Email is Succes";
        } catch (MessagingException e) {
            e.printStackTrace();
            return "Sending Email is Failed";
        }
    }
}
