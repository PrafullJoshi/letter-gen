package in.co.parna.lettergen.service;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import in.co.parna.lettergen.dto.GmailCredentials;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Iterator;

@Service
public class GmailLetterSenderService {

    @Autowired
    private GmailService gmailService;

    @Value("${gmail.senderGmailId}")
    private String senderGmailId;

    @Value("${gmail.secret.clientId}")
    private String clientId;

    @Value("${gmail.secret.clientSecret}")
    private String clientSecret;

    @Value("${gmail.secret.accessToken}")
    private String accessToken;

    @Value("${gmail.secret.refreshToken}")
    private String refreshToken;


    private String RECIPIENT = "PRAFULL.JOSHI@GMAIL.COM";

    @Value("${letter.sender.testMode}")
    private boolean testMode;

    @Value("${letter.sender.membersListFilePath}")
    private String MEMBERS_EXCEL_FILE_PATH;


    public void sendIndividualLettersWithAttachment() {
        try {
            gmailService.setHttpTransport(GoogleNetHttpTransport.newTrustedTransport());
            gmailService.setGmailCredentials(GmailCredentials.builder()
                    .userEmail(senderGmailId)
                    .clientId(clientId)
                    .clientSecret(clientSecret)
                    .accessToken(accessToken)
                    .refreshToken(refreshToken)
                    .build());


            String to = RECIPIENT; // list of recipient email addresses
            String subject = "FY19-20 annual maintenance demand";

            String body = "         Please find enclosed the FY19-20 annual maintenance demand.  Kindly pay the total amount payable by 31-Dec-2019 to avoid late fee charges."
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
            int threshold = testMode ? 2 : 200;
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

                to = testMode ? "prafull.joshi@gmail.com" : emails;
                String finalBody = "<br>" +greetings + name + " (" + flatNo + "),<br><br>" + body;
                try {
//                    sendWithAttachment(to, subject, finalBody, flatNo.trim());
                    gmailService.sendMessage(to, subject, finalBody, flatNo.trim());
                    System.out.println("Mail successfully sent to - " + flatNo + " " + to);
                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Mail sending to - " + flatNo + " " + to + " failed!");
                }
            }




        } catch (GeneralSecurityException | IOException e) {
            e.printStackTrace();
        }
    }
}
