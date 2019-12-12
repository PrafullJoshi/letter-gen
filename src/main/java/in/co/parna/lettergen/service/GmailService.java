package in.co.parna.lettergen.service;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.Base64;
import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.model.Message;
import in.co.parna.lettergen.dto.GmailCredentials;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Properties;

@Service
@Slf4j
public final class GmailService {

    @Value("${gmail.applicationName}")
    private String APPLICATION_NAME;

    private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();

    private HttpTransport httpTransport;

    private GmailCredentials gmailCredentials;

    @Value("${letter.sender.attachmentNameShownInMail}")
    private String ATTACHMENT_FILE_NAME_1;

    @Value("${letter.sender.attachmentFilePath}")
    private String ATTACHMENT_1;

    public GmailService() {

    }

    public HttpTransport getHttpTransport() {
        return httpTransport;
    }

    public void setHttpTransport(HttpTransport httpTransport) {
        this.httpTransport = httpTransport;
    }

    public void setGmailCredentials(GmailCredentials gmailCredentials) {
        this.gmailCredentials = gmailCredentials;
    }

    public boolean sendMessage(String recipientAddresses, String subject, String body, String flatNumber) throws MessagingException,
            IOException {
        Message message = createMessageWithEmail(
                createEmail(recipientAddresses, gmailCredentials.getUserEmail(), subject, body, flatNumber));

        return createGmail().users()
                .messages()
                .send(gmailCredentials.getUserEmail(), message)
                .execute()
                .getLabelIds().contains("SENT");
    }

    private Gmail createGmail() {
        Credential credential = authorize();
        return new Gmail.Builder(httpTransport, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    private MimeMessage createEmail(String recipientAddresses, String from, String subject, String bodyText, String flatNumber) throws MessagingException {
        MimeMessage email = new MimeMessage(Session.getDefaultInstance(new Properties(), null));
        email.setFrom(new InternetAddress(from));
        email.addRecipients(javax.mail.Message.RecipientType.TO, InternetAddress.parse(recipientAddresses));
        email.setSubject(subject);
        email.setText(bodyText, "utf-8", "html"); // TODO "utf-8", "html"

        Multipart multipart = new MimeMultipart();
        BodyPart messageBodyPartAttachment = new MimeBodyPart();

        String attachmentPath = ATTACHMENT_1.replace("$", flatNumber);
        DataSource source = new FileDataSource(attachmentPath);

        messageBodyPartAttachment.setDataHandler(new DataHandler(source));

        String attachmentName = ATTACHMENT_FILE_NAME_1.replace("$", flatNumber);
        messageBodyPartAttachment.setFileName(attachmentName);
        multipart.addBodyPart(messageBodyPartAttachment);

        BodyPart messageBodyPart = new MimeBodyPart();
        messageBodyPart.setContent(bodyText, "text/html");
        multipart.addBodyPart(messageBodyPart);

        email.setContent(multipart);
        return email;
    }

    private Message createMessageWithEmail(MimeMessage emailContent) throws MessagingException, IOException {
        ByteArrayOutputStream buffer = new ByteArrayOutputStream();
        emailContent.writeTo(buffer);

        return new Message()
                .setRaw(Base64.encodeBase64URLSafeString(buffer.toByteArray()));
    }

    private Credential authorize() {
        return new GoogleCredential.Builder()
                .setTransport(httpTransport)
                .setJsonFactory(JSON_FACTORY)
                .setClientSecrets(gmailCredentials.getClientId(), gmailCredentials.getClientSecret())
                .build()
                .setAccessToken(gmailCredentials.getAccessToken())
                .setRefreshToken(gmailCredentials.getRefreshToken());
    }
}
