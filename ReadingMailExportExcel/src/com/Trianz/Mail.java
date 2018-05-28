package com.Trianz;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Properties;
import java.util.Scanner;

import javax.mail.FetchProfile;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;


 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


 
public class Mail {
	private static final String FILE_NAME = "C:\\Users\\"+System.getProperty("user.name")+"\\Desktop\\MailListex.xlsx";
    public static void main(String[] args) throws IOException, MessagingException{
    	System.out.println("Welcome to mail reading application");
    	System.out.println(System.getProperty("user.name"));
    	
    	Mail mail = new Mail();
        Message[] msg=mail.read();
        mail.exportExcel(msg);
        
 }
    	 
    	
    public void exportExcel(Message[] msg) throws MessagingException {
    	System.out.println("entered into exportExcel()");
    	XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Mails Reading");
      
        Object[][] datatypes=new Object[msg.length+1][6];
     
        datatypes[0][0]="From";
        datatypes[0][1]="To";
        datatypes[0][2]="Subject";
        datatypes[0][3]="Received date";
      //  datatypes[0][4]="Sent date";
        System.out.println(msg.length+"hi ");
        for(int i=0;i<msg.length;i++) {
        	datatypes[i+1][0] = msg[i].getFrom()[0].toString();
        //	datatypes[i+1][1] =  msg[i].getAllRecipients()[0].toString();
          	datatypes[i+1][2] =msg[i].getSubject();
          	datatypes[i+1][3] = msg[i].getReceivedDate().toString();
        	// datatypes[i][4]=msg[i].getSentDate().toString();
        	 
        	
        }

        int rowNum = 0;
        System.out.println("Creating excel");
        XSSFFont defaultFont= workbook.createFont();
        
        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            
            int colNum = 0;
           
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                    
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
              
                
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
           // e.printStackTrace();
            System.out.println("Please close the Excel File and run the application again.");
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    
    }
 
    
    public Message[] read()  {
    	System.out.println("inside Read()");
    	 Message[] messages = null;
 
        //Properties props = new Properties();
 
        try {
        	       			
        			Properties props = new Properties();
        			props.put("mail.smtp.starttls.enable", "true");
        			props.put("mail.smtp.auth", "true");
        			//props.put("mail.imap.usesocketchannels", "true");
        			props.put("mail.smtp.host", "m.outlook.com");//m.outlook.com//smtp.gmail.com
        			props.put("mail.smtp.port", "993");//993//465
        			props.put("mail.smtp.socketFactory.port", "993");//993//465
        			props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
        			props.setProperty("mail.imaps.usesocketchannels", "true");
        	
        	
            //props.load(new FileInputStream(new File("C:\\smtp.properties")));
            Session session = Session.getDefaultInstance(props, null);
 
            Store store = session.getStore("imaps");
            store.connect("m.outlook.com", "pavan.gurram@trianz.com", "Trianz?$69910");//smtp.gmail.com//m.outlook.com
            
            Folder inbox = store.getFolder("INBOX");
           

           
            
            inbox.open(Folder.READ_ONLY);
            int messageCount = inbox.getMessageCount();
 
            System.out.println("Total Messages:- " + messageCount);
          
             messages = inbox.getMessages();
            System.out.println("-----message length---"+ messages.length);
 
            for (int i =100; i>0; i--) {
                System.out.println("Mail Subject:- " + messages[i].getSubject());
                System.out.println("sent date:- " + messages[i].getSentDate());
                System.out.println("receipent:- " + messages[i].getAllRecipients()[0]);//Use StringTokenizer to send email addr.
                System.out.println("received date:- " + messages[i].getReceivedDate());
                System.out.println("from:- " + messages[i].getFrom()[0].toString());
           
              //  System.out.println("Mail content:- " + messages[i].getContent().toString());
                System.out.println("------------------------------"+(i+1));
            }
           inbox.close(true);
            store.close();
            return messages;
 
        } catch (Exception e) {
        	System.out.println("entered into read() catch");
        	return messages;
          //  e.printStackTrace();
        }
       // return null;
    }

	

 
}