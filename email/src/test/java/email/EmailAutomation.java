package email;


	import java.io.FileOutputStream;
import java.io.IOException;
	import java.util.Properties;

	import javax.mail.*;
	import javax.mail.internet.MimeMultipart;

	import org.apache.poi.ss.usermodel.*;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	public class EmailAutomation {
	    public static void main(String[] args) {
	        // Email configuration
	        String host = "imap.gmail.com";
	        String username = "neil@growpital";		
	        String password = "Liverpool@01";

	        // Connect to the email server and retrieve emails
	        Properties properties = new Properties();	
	        properties.put("mail.store.protocol", "imaps");
	        properties.put("mail.imap.connectiontimeout", "5000"); // Set connection timeout to 5 seconds
	        properties.put("mail.imap.socketFactory.timeout", "5000"); // Set socket timeout to 5 seconds


	        try {
	            Session session = Session.getDefaultInstance(properties);
	            Store store = session.getStore("imaps");
	            store.connect(host, username, password);

	            Folder inbox = store.getFolder("INBOX");
	            inbox.open(Folder.READ_ONLY);

	            Message[] messages = inbox.getMessages();

	            // Process each email
	            for (Message message : messages) {
	                String subject = message.getSubject();
	                System.out.println("Subject: " + subject);

	                // You can add more conditions to filter emails if needed
	                if (subject.contains("Lambdatest > Trial Activation")) {
	                    // Extract content from the email
	                    String emailContent = extractEmailContent(message);

	                    // Copy and paste content to an Excel file (You need to implement this part)
	                    writeToExcel(emailContent);
	                }
	            }

	            inbox.close(false);
	            store.close();

	        } catch (MessagingException | IOException e) {
	            e.printStackTrace();
	        }
	    }

	    private static String extractEmailContent(Message message) throws IOException, MessagingException {
	    	
	    
	        Object content = message.getContent();

	        if (content instanceof MimeMultipart) {
	            MimeMultipart mimeMultipart = (MimeMultipart) content;

	            // You may need to modify this loop to select the desired part
	            for (int i = 0; i < mimeMultipart.getCount(); i++) {
	                BodyPart bodyPart = mimeMultipart.getBodyPart(i);
	                if (bodyPart.isMimeType("text/plain")) {
	                    return bodyPart.getContent().toString();
	                }
	            }
	        }
	        return "";
	    }
	    
	    private static void writeToExcel(String content) {
	        // Create an Excel workbook and write the content to it
	        Workbook workbook = new XSSFWorkbook();
	        Sheet sheet = workbook.createSheet("Email Content");
	        Row row = sheet.createRow(0);
	        Cell cell = row.createCell(0);
	        cell.setCellValue(content);

	        // You can customize the Excel file writing logic here

	        try {
	            FileOutputStream outputStream = new FileOutputStream("email_content.xlsx");
	            workbook.write(outputStream);
	            outputStream.close();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	}
