
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.Properties;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.activation.DataHandler;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.search.FromTerm;
import javax.mail.search.SearchTerm;

import org.apache.commons.io.FilenameUtils;

import com.sun.mail.util.MailSSLSocketFactory;

public class EmailSearcher {

	public static void readOutlookEmail(String senderName, String fileDownloadPath, String host, String port,
			String userEmail, String password) throws GeneralSecurityException {
		Properties properties = new Properties();

		// server setting
		properties.put("mail.imap.host", host);
		properties.put("mail.imap.port", port);

		// SSL setting
		MailSSLSocketFactory socketFactory = new MailSSLSocketFactory();
		socketFactory.setTrustAllHosts(true);
		properties.put("mail.imap.ssl.socketFactory", socketFactory);

		properties.setProperty("mail.imap.socketFactory.class", "javax.net.ssl.SSLSocketFactory");
		properties.setProperty("mail.imap.socketFactory.fallback", "false");
		properties.setProperty("mail.imap.socketFactory.port", String.valueOf(port));
		properties.setProperty("mail.imaps.auth.plain.disable", "true");
		properties.setProperty("mail.imaps.auth.ntlm.disable", "true");
		properties.setProperty("mail.store.protocol", "imaps");

		Session session = Session.getInstance(properties);
		session.setDebug(true);

		try {
			// connects to the message store
			Store store = session.getStore("imaps");
			store.connect(host, userEmail, password);

			// opens the inbox folder
			Folder folderInbox = store.getFolder("INBOX");
			folderInbox.open(Folder.READ_ONLY);

			SearchTerm searchCondition = new FromTerm(new InternetAddress(senderName));

			// performs search through the folder
			Message[] foundMessages = folderInbox.search(searchCondition);

			// To get email from latest to older order
			Arrays.sort(foundMessages, (m1, m2) -> {
				try {
					return m2.getSentDate().compareTo(m1.getSentDate());
				} catch (MessagingException e) {
					throw new RuntimeException(e);
				}
			});

			for (int i = 0; i < foundMessages.length; i++) {
				Message message = foundMessages[i];
				String subject = message.getSubject();
				System.out.println("Found message #" + i + ": with Subject -> " + subject);

				if (!(message.getContent() instanceof String)) {
					Multipart multipart = (Multipart) message.getContent();
					for (int j = 0; j < multipart.getCount(); j++) {
						MimeBodyPart bodyPart = (MimeBodyPart) multipart.getBodyPart(j);
						String disposition = bodyPart.getDisposition();
						if (disposition != null && Part.ATTACHMENT.equalsIgnoreCase(disposition)) {
							DataHandler handler = bodyPart.getDataHandler();
							String fileExt = FilenameUtils.getExtension(handler.getName());
							String fullPath = fileDownloadPath + File.separator + bodyPart.getFileName();

							// System.out.println("content-type: " + bodyPart.getContentType());
							if (fileExt.equals("txt") || bodyPart.isMimeType("text/plain")) {
								((MimeBodyPart) bodyPart).saveFile(fullPath);
							} else if (fileExt.equals("zip") || bodyPart.isMimeType("APPLICATION/X-ZIP-COMPRESSED")) {
								((MimeBodyPart) bodyPart).saveFile(fullPath);
								unZip(fullPath, fileDownloadPath);
							}
						}
					}
				}
			}

			// disconnect
			folderInbox.close(false);
			store.close();
		} catch (NoSuchProviderException ex) {
			System.out.println("No provider.");
			ex.printStackTrace();
		} catch (MessagingException ex) {
			System.out.println("Could not connect to the message store.");
			ex.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void unZip(String zipFile, String outputFolder) {
		byte[] buffer = new byte[1024];

		try {
			// create output directory is not exists
			File folder = new File(outputFolder);
			if (!folder.exists()) {
				folder.mkdir();
			}

			// get the zip file content
			ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile));

			// get the zipped file list entry
			ZipEntry ze = zis.getNextEntry();

			while (ze != null) {
				String fileName = ze.getName();
				File newFile = new File(outputFolder + File.separator + fileName);
				// System.out.println("file unzip : " + newFile.getAbsoluteFile());

				// create all non exists folders
				// else you will hit FileNotFoundException for compressed folder
				new File(newFile.getParent()).mkdirs();

				FileOutputStream fos = new FileOutputStream(newFile);

				int len;
				while ((len = zis.read(buffer)) > 0) {
					fos.write(buffer, 0, len);
				}

				fos.close();
				ze = zis.getNextEntry();
			}

			zis.closeEntry();
			zis.close();

			// To delete zip file after extract content
			// File file = new File(zipFile);
			// file.delete();
			// System.out.println("Done");
		} catch (IOException ex) {
			ex.printStackTrace();
		}
	}

	public static void main(String[] args) throws GeneralSecurityException {
		String host = "outlook.office365.com"; // Validate if you are using the same or not
		String userEmail = "userEmail"; // Pass valid user email address
		String password = "password"; // Pass valid user password
		String port = "993";
		String fileDownloadPath = "E:\\client\\download\\attachments"; // Pass path to download file from attachment
		// email address to fetch all the email's from specific sender
		String senderName = "senderName"; // Pass valid sender
		readOutlookEmail(senderName, fileDownloadPath, host, port, userEmail, password);
	}

}
