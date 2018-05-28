package com.Trianz;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;

import javax.mail.Address;
import javax.mail.FetchProfile;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Session;
import javax.mail.Store;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingMails {

		private static final String FILE_NAME = "C:\\Users\\"+System.getProperty("user.name")+"\\Desktop\\MailList.xlsx";
	    public static void main(String[] args) throws IOException, MessagingException{
	    	System.out.println("Welcome to mail reading application");
	    	System.out.println(System.getProperty("user.name"));
	    	
	    	ReadingMails mail = new ReadingMails();
	        Message[] msg=mail.read();
	        mail.exportExcel(msg);
	        }
	    
	    public void exportExcel(Message[] msg) throws MessagingException {
	    	System.out.println("entered into exportExcel()");
	    	XSSFWorkbook workbook = new XSSFWorkbook();
	        XSSFSheet sheet = workbook.createSheet("Mails Reading");
	      
	        Object[][] datatypes=new Object[msg.length+1][6];
	     
	        datatypes[0][0]="From";
	        datatypes[0][1]="Subject";
	        datatypes[0][2]="Received date";
	        datatypes[0][3]="Sent date";
	      //  datatypes[0][4]="Sent date";
	        System.out.println(msg.length+"hi ");
	        SimpleDateFormat df = new SimpleDateFormat("dd-MM-yyyy");
	        for(int i=0,j=msg.length;i<msg.length;i++,j--) {
	        	
	        	Date receivedDate = msg[i].getReceivedDate();
	                Date sentDate = msg[i].getSentDate(); 

	                  
	              //  SimpleDateFormat df = new SimpleDateFormat("dd-MM-yyyy");
	        
	        	datatypes[j][0] = msg[i].getFrom()[0].toString();
	          	datatypes[j][1] =msg[i].getSubject();
	          	datatypes[j][2] =df.format(receivedDate);
	          	datatypes[j][3] =df.format(sentDate);
	        	// datatypes[i][4]=msg[i].getSentDate().toString();
	        	 
	        	
	        }

	        int rowNum = 0;
	        System.out.println("Creating excel");
	      //  XSSFFont defaultFont= workbook.createFont();
	        
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
	    	
	    	System.out.println("Inside MailReader()...");
	        final String SSL_FACTORY = "javax.net.ssl.SSLSocketFactory";
	        
	        Properties props = System.getProperties();
	        // Set manual Properties
	        props.setProperty("mail.imaps.socketFactory.class", SSL_FACTORY);
	        props.setProperty("mail.imaps.socketFactory.fallback", "false");
	        props.setProperty("mail.imaps.port", "993");
	        props.setProperty("mail.imaps.socketFactory.port", "993");
	        props.put("mail.imaps.host", "imap-mail.outlook.com");


	        try {

	            Session session = Session.getDefaultInstance(System.getProperties(), null);
	            Store store = session.getStore("imaps");

	            store.connect("imap-mail.outlook.com", 993, "pavan.gurram@trianz.com", "Trianz?$69910");
	            Folder inbox = store.getFolder("INBOX");

	            inbox.open(Folder.READ_ONLY);
	            
	            Message[] messages=inbox.getMessages();

	           FetchProfile fp = new FetchProfile();
	            fp.add(FetchProfile.Item.ENVELOPE);

	            inbox.fetch(messages, fp);

	            try {

	               // printAllMessages(messages);
	                Address[] a;

	                // FROM
	                for (int i = messages.length-1; i>0; i--) {
	                	
	                	 if ((a = messages[i].getFrom()) != null) {
	                		 
	 	                    for (int j = 0; j < a.length; j++) {
	 	                        System.out.println("From : " + a[j].toString());
	 	                    }
	 	                }
	 	           

	 	                String subject = messages[i].getSubject();

	 	                Date receivedDate = messages[i].getReceivedDate();
	 	                Date sentDate = messages[i].getSentDate(); 

	 	                SimpleDateFormat df = new SimpleDateFormat("dd-MM-yyyy");

	 	                System.out.println("Subject : " + subject);

	 	                if (receivedDate != null) {
	 	                    System.out.println("ReceivedDate: " + df.format(receivedDate));
	 	                }

	 	                System.out.println("SentDate : " + df.format(sentDate));
	                }
	               

	                inbox.close(true);
	                store.close();
	                return messages;

	            } catch (Exception ex) {
	                System.out.println("Exception arise at the time of read mail");
	                ex.printStackTrace();
	            }

	        } catch (MessagingException e) {
	            System.out.println("Exception while connecting to server: " + e.getLocalizedMessage());
	            e.printStackTrace();
	            System.exit(2);
	        }
			return null;

	    }

}
