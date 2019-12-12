package in.co.parna.lettergen.service;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Properties;

@Service
public class LetterSenderService {

    @Value("${letter.generator.inputFilePathname}")
    String inputFilePathname;

    @Value("${letter.generator.logoPath}")
    String logoPath;

    @Value("${letter.generator.document.count.threshold}")
    int threshold;

    private static String OUTPUT_FILE = "output/#_Madhukosh_Maintainance.pdf";

    @Value("${letter.sender.membersListFilePath}")
    String MEMBERS_EXCEL_FILE_PATH;

    @Value("${letter.sender.attachmentNameShownInMail}")
    private String ATTACHMENT_FILE_NAME_1;

    @Value("${letter.sender.attachmentFilePath}")
    private String ATTACHMENT_1;


    private static String RECIPIENT = "PRAFULL.JOSHI@GMAIL.COM";
//    private static String FROM = "madhukoshfg@gmail.com";
//    private static String PASSWORD = "Bhagwati2mail%";
    private static String FROM = "prafull.joshi@gmail.com";
    private static String PASSWORD = "mh12gw6779";

    public void sendIndividualLettersWithAttachment() throws IOException {
        sendFromExcelList();
    }

    private void sendFromExcelList() throws IOException {
        String to = RECIPIENT; // list of recipient email addresses

        String subject = "FY19-20 annual maintenance demand";

//        String body = "<p class=\"MsoNormal\" style=\"color: #222222; font-family: Arial, Helvetica, sans-serif; font-size: small; background-color: #ffffff\"><span style=\"font-size: 10pt; line-height: 14.2667px\">Please find enclosed the&nbsp;<b>FY19-20: Annual Maintenance Charges Demand&nbsp;</b>letter<b>&nbsp;</b>for the period&nbsp;<b>April-2019 to September-2019.</b>.&nbsp; Please&nbsp;<b>read complete letter carefully</b>&nbsp;to understand all the details.&nbsp; Please&nbsp;<b>pay the \"Total Amount Payable\"</b>&nbsp;as per the schedule provided in the letter.<u></u><u></u></span></p><br>" +
//                "<p class=\"MsoNormal\" style=\"color: #222222; font-family: Arial, Helvetica, sans-serif; font-size: small; background-color: #ffffff\"><span style=\"font-size: 10pt; line-height: 14.2667px\">The committee requests all members to avail the prompt discount benefit, in case it is not possible please ensure that you pay before the late due date to avoid interest charges penalty.</span></p>"

        String body = "<br><br>"
                + "Please find enclosed the FY19-20 annual maintenance demand.  Kindly pay the total amount payable by 31-Dec-2019 to avoid late fee charges."
                + "<br><br>"
                + "Thank you,<br>" +
                "Madhukosh FG Committee<br>" +
                "<strong>Aditya Barde (Secretary)</strong><br>" +
                "<strong>Chandrakant Mandape (Treasurer)</strong><br>" +
                "<strong>Shrinivas Dombe (President)</strong>";

        FileInputStream excelFile = new FileInputStream(new File(MEMBERS_EXCEL_FILE_PATH));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();
        int counter = 0;
        int threshold = 2;
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            Iterator<org.apache.poi.ss.usermodel.Cell> cellIterator = currentRow.iterator();
            if(currentRow.getRowNum() == 0) {
                continue;
            }

            if(counter == threshold) {
                break;
            }

            counter++;
            String greetings = "Dear ";
            String flatNo = "";
            String name = "";
            String emails = "";
            while (cellIterator.hasNext()) {
                org.apache.poi.ss.usermodel.Cell currentCell = cellIterator.next();
                switch(currentCell.getColumnIndex()) {
                    case 0 :
                        flatNo = currentCell.getStringCellValue();
                        break;
                    case 1 :
                        name = currentCell.getStringCellValue();
                        break;
                    case 2 :
                        emails = currentCell.getStringCellValue();
                        break;
                }
            }

            to = emails;
            String finalBody = "<br>" +greetings + name + " (" + flatNo + "),<br><br>" + body;
            try {
                sendWithAttachment(to, subject, finalBody, flatNo.trim());
                System.out.println("Mail successfully sent to - " + flatNo + " " + to);
            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("Mail sending to - " + flatNo + " " + to + " failed!");
            }
        }
    }

    private void sendWithAttachment(String to, String subject,
                                           String body, String flatNumber) throws MessagingException {
        Properties props = System.getProperties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        String host = "smtp.gmail.com";
        props.put("mail.smtp.host", host);
        props.put("mail.smtp.port", "587");

        props.put("mail.smtp.user", FROM);
        props.put("mail.smtp.password", PASSWORD);
        Session session = Session.getInstance(props, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(FROM, PASSWORD);
            }
        });
        MimeMessage message = new MimeMessage(session);
        message.setFrom(new InternetAddress(FROM));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(RECIPIENT)); // Actual Receipients
        message.setSubject(subject);
        message.setText(body, "utf-8", "html");

        // message.saveChanges();


        Multipart multipart = new MimeMultipart();
        BodyPart messageBodyPartAttachment = new MimeBodyPart();

        String attachmentPath = ATTACHMENT_1.replace("$", flatNumber);
        DataSource source = new FileDataSource(attachmentPath);

        messageBodyPartAttachment.setDataHandler(new DataHandler(source));

        String attachmentName = ATTACHMENT_FILE_NAME_1.replace("$", flatNumber);
        messageBodyPartAttachment.setFileName(attachmentName);
        multipart.addBodyPart(messageBodyPartAttachment);

        /*messageBodyPartAttachment = new MimeBodyPart();
        source = new FileDataSource(ATTACHMENT_2);
        messageBodyPartAttachment.setDataHandler(new DataHandler(source));
        messageBodyPartAttachment.setFileName(ATTACHMENT_FILE_NAME_2);
        multipart.addBodyPart(messageBodyPartAttachment);*/

        BodyPart messageBodyPart = new MimeBodyPart();
        messageBodyPart.setContent(body, "text/html");
        multipart.addBodyPart(messageBodyPart);




        // Send the complete message parts
        message.setContent(multipart);

        Transport.send(message);
    }
}
