
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.FolderId;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.property.complex.Mailbox;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

public class ExchangeEmailTest {

	private static ExchangeService service;
	private static Integer NUMBER_EMAILS_FETCH = 5;

	static {
		try {
			service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
			service.setUrl(new URI("https://outlook.office365.com/EWS/Exchange.asmx")); // Exchange URL of Outlook Server
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public ExchangeEmailTest() {
		ExchangeCredentials credentials = new WebCredentials("CLIENT_EMAIL", "CLIENT_PASSWORD");
		service.setCredentials(credentials);
	}

	public List readEmails(List<String> emailList, String storeEmailPath) {
		List msgDataList = new ArrayList<>();

		for (String mailBoxEmail : emailList) {
			try {
				Folder folder = Folder.bind(service, new FolderId(WellKnownFolderName.Inbox, new Mailbox(mailBoxEmail)));

				countTotalEmails(folder);
				FindItemsResults results = service.findItems(folder.getId(), new ItemView(NUMBER_EMAILS_FETCH));
				int i = 1;
				Iterator<Item> items = results.iterator();

				while (items.hasNext()) {
					Item item = items.next();
					Map messageData = new HashMap();
					messageData = readEmailItem(item.getId());
					System.out.println("\nEmails #" + (i++) + ":");
					System.out.println("Subject : " + messageData.get("subject"));
					System.out.println("Sender Email ID : " + messageData.get("fromAddress"));
					System.out.println("Sender Name : " + messageData.get("senderName"));
					System.out.println("Received Date : " + messageData.get("ReceivedDate"));

					EmailMessage emailMessage = (EmailMessage) messageData.get("emailMessage");

					checkAttachments(emailMessage);
					copyMailToFolder(storeEmailPath, emailMessage);

					msgDataList.add(messageData);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return msgDataList;
	}

	private void countTotalEmails(Folder folder) throws ServiceLocalException {
		Integer totalCount = folder.getTotalCount();
		System.out.println("Total Mails in Inbox ===>>>" + totalCount);

		if (totalCount < NUMBER_EMAILS_FETCH) {
			NUMBER_EMAILS_FETCH = totalCount;
		}
	}

	public Map readEmailItem(ItemId itemId) {
		Map messageData = new HashMap<>();
		try {
			PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties);
			propSet.add(ItemSchema.MimeContent);

			Item itm = Item.bind(service, itemId, PropertySet.FirstClassProperties);
			EmailMessage emailMessage = EmailMessage.bind(service, itm.getId(), propSet);
			messageData.put("emailItemId", emailMessage.getId().toString());
			messageData.put("subject", emailMessage.getSubject().toString());
			messageData.put("fromAddress", emailMessage.getFrom().getAddress().toString());
			messageData.put("senderName", emailMessage.getFrom().getName().toString());
			Date dateTimeCreated = emailMessage.getDateTimeCreated();
			messageData.put("SendDate", dateTimeCreated.toString());
			Date dateTimeReceived = emailMessage.getDateTimeReceived();
			messageData.put("ReceivedDate", dateTimeReceived.toString());
			messageData.put("Size", String.valueOf(emailMessage.getSize()));
			messageData.put("emailBody", emailMessage.getBody().toString());
			messageData.put("emailMessage", emailMessage);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return messageData;
	}

	private void checkAttachments(EmailMessage emailMessage) throws ServiceLocalException {

		// To check and read attachment name
		if (emailMessage.getHasAttachments()) {
			AttachmentCollection attachmentCollection = emailMessage.getAttachments();
			for (Attachment attachment : attachmentCollection.getItems()) {
				if (!attachment.getContentType().equals("image/png")
						&& !attachment.getContentType().equals("image/jpeg")
						&& !attachment.getContentType().equals("image/jpg")
						&& !attachment.getContentType().equals("image/gif")) {
					System.out.println("Attachment Name ==>> " + attachment.getName());
				}
			}
		}
	}

	private void copyMailToFolder(String storeEmailPath, EmailMessage emailMessage)
			throws ServiceLocalException, IOException {

		// To save email into destination folder
		File toSave = new File(storeEmailPath + emailMessage.getSubject() + ".eml");
		FileUtils.writeByteArrayToFile(toSave, emailMessage.getMimeContent().getContent());
	}

	public static void main(String[] args) {
		String storeEmailPath = "PATH_OF_FILE_SYSTEM_TO_COPY_EMAIL";

		ExchangeEmailTest exchangeEmailTest = new ExchangeEmailTest();

		List<String> emailList = new ArrayList<String>();
		emailList.add("Mailbox_EMAIL_1");
		emailList.add("Mailbox_EMAIL_2");
		emailList.add("Mailbox_EMAIL_3");

		exchangeEmailTest.readEmails(emailList, storeEmailPath);
	}
}
