
import java.io.File;
import java.net.URI;
import java.util.Iterator;

import org.apache.commons.io.FileUtils;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

public class EWSReadEmails {

	private static ExchangeService service;
	private static Integer NUMBER_EMAILS_FETCH = 5;

	public void readEmails(String outlookEmailAddress, String outlookEmailPassword, String exchangeUrl,
			String storeEmailPath) {

		try {
			PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties);
			propSet.add(ItemSchema.MimeContent);

			service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
			ExchangeCredentials credentials = new WebCredentials(outlookEmailAddress, outlookEmailPassword);
			service.setCredentials(credentials);
			service.setUrl(new URI(exchangeUrl));

			Folder folder = Folder.bind(service, WellKnownFolderName.Inbox);
			Integer totalCount = folder.getTotalCount();
			System.out.println("Total Mails in Inbox ===>>>" + totalCount);

			if (totalCount < NUMBER_EMAILS_FETCH) {
				NUMBER_EMAILS_FETCH = totalCount;
			}
			FindItemsResults<Item> results = service.findItems(folder.getId(), new ItemView(NUMBER_EMAILS_FETCH));
			Iterator<Item> items = results.iterator();
			int i = 1;

			while (items.hasNext()) {
				Item item = items.next();
				EmailMessage emailMessage = EmailMessage.bind(service, item.getId(), propSet);
				System.out.println("Email Item No ================>> " + i++);
				System.out.println("Subject ==========>> " + emailMessage.getSubject());
				System.out.println("Sender Email ID ==>> " + emailMessage.getFrom().getAddress());
				System.out.println("Sender Name ======>> " + emailMessage.getFrom().getName());
				System.out.println("Received Date ====>> " + emailMessage.getDateTimeReceived());
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

				File toSave = new File(storeEmailPath + emailMessage.getSubject() + ".eml");
				FileUtils.writeByteArrayToFile(toSave, emailMessage.getMimeContent().getContent());
			}
		} catch (Exception e) {
			System.out.println("Exception in readEmails ==>>" + e);
		}
	}

	public static void main(String[] args) {
		String outlookEmailAddress = "EMAIL_ADDRESS";
		String outlookEmailPassword = "PASSWORD";
		String exchangeUrl = "https://outlook.office365.com/EWS/Exchange.asmx"; // Depends on your exchange server
		String storeEmailPath = "FOLDER_LOCATION_TO_STORE_EMAIL";

		EWSReadEmails ewsReadEmails = new EWSReadEmails();
		ewsReadEmails.readEmails(outlookEmailAddress, outlookEmailPassword, exchangeUrl, storeEmailPath);
	}
}
