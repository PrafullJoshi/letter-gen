package in.co.parna.lettergen.api;

import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import in.co.parna.lettergen.dto.GmailCredentials;
import in.co.parna.lettergen.service.GmailLetterSenderService;
import in.co.parna.lettergen.service.GmailService;
import in.co.parna.lettergen.service.LetterSenderService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.mail.MessagingException;
import java.io.IOException;
import java.security.GeneralSecurityException;

@Api(value = "Letter Sender", tags = "Letter Sender")
@RestController("api/v1/letter-sender")
public class LetterSenderController {

    @Autowired
    private LetterSenderService letterSenderService;

    @Autowired
    private GmailLetterSenderService gmailLetterSenderService;

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

    @ApiOperation(value = "Letter Sender with provided information", tags = "Letter Sender")
    @RequestMapping(value = "/bulk", method = RequestMethod.GET)
    public void sendBulkLetters() throws IOException {

        gmailLetterSenderService.sendIndividualLettersWithAttachment();
    }

    @ApiOperation(value = "Letter Sender with provided information", tags = "Letter Sender")
    @RequestMapping(method = RequestMethod.GET)
    public void sendLetters() throws IOException {

        letterSenderService.sendIndividualLettersWithAttachment();
    }

    @ApiOperation(value = "Test Ping Mail", tags = "Letter Sender")
    @RequestMapping(value = "/gmail-api/test", method = RequestMethod.GET)
    public void testMail() throws IOException {
        try {
            gmailService.setHttpTransport(GoogleNetHttpTransport.newTrustedTransport());
            gmailService.setGmailCredentials(GmailCredentials.builder()
                    .userEmail(senderGmailId)
                    .clientId(clientId)
                    .clientSecret(clientSecret)
                    .accessToken(accessToken)
                    .refreshToken(refreshToken)
                    .build());

            gmailService.sendMessage("prafullpjoshi@yahoo.co.in", "Subject", "body text", "G-1005");
        } catch (GeneralSecurityException | IOException | MessagingException e) {
            e.printStackTrace();
        }
    }
}
