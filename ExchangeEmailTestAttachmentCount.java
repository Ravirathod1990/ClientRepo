
import java.net.URI;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

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
import microsoft.exchange.webservices.data.property.complex.Mailbox;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

public class ExchangeEmailTestAttachmentCount {

	private static ExchangeService service;
	private static Integer NUMBER_EMAILS_FETCH = 5;

	static {
		try {
			service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
			service.setUrl(new URI("https://outlook.office365.com/EWS/Exchange.asmx"));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public ExchangeEmailTestAttachmentCount() {
		ExchangeCredentials credentials = new WebCredentials("EMAIL_ID", "PASSWORD");
		service.setCredentials(credentials);
	}

	public void readEmails(List<String> emailList, String storeEmailPath) {
		for (String mailBoxEmail : emailList) {
			try {
				Folder folder = Folder.bind(service,
						new FolderId(WellKnownFolderName.Inbox, new Mailbox(mailBoxEmail)));

				countTotalEmails(folder);
				FindItemsResults results = service.findItems(folder.getId(), new ItemView(NUMBER_EMAILS_FETCH));
				int i = 1;
				Iterator<Item> items = results.iterator();

				while (items.hasNext()) {
					Item item = items.next();
					PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties);
					propSet.add(ItemSchema.MimeContent);

					Item itm = Item.bind(service, item.getId(), PropertySet.FirstClassProperties);
					EmailMessage emailMessage = EmailMessage.bind(service, itm.getId(), propSet);
					System.out.println("Subject : " + emailMessage.getSubject());

					// Requirement 3
					// Check and count attachments
					int totalAttachments = countAttachments(emailMessage);
					System.out.println("Total Attachments ==>> " + totalAttachments);
					if (totalAttachments > 0) {
						List<String> attachmentsName = getAttachmentsName(emailMessage);
						System.out.println("Attachment Name ==>> " + attachmentsName);
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	private void countTotalEmails(Folder folder) throws ServiceLocalException {
		Integer totalCount = folder.getTotalCount();
		System.out.println("Total Mails in Inbox ===>>> " + totalCount);

		if (totalCount < NUMBER_EMAILS_FETCH) {
			NUMBER_EMAILS_FETCH = totalCount;
		}
	}

	// Requirement 3
	// Count attachments
	private int countAttachments(EmailMessage emailMessage) throws ServiceLocalException {
		int totalAttachment = 0;

		// To check and read attachment name
		if (emailMessage.getHasAttachments()) {
			totalAttachment = emailMessage.getAttachments().getCount();
			AttachmentCollection attachmentCollection = emailMessage.getAttachments();
			for (Attachment attachment : attachmentCollection.getItems()) {
				if (attachment.getContentType().equals("image/png") 
						|| attachment.getContentType().equals("image/jpeg")
						|| attachment.getContentType().equals("image/jpg")
						|| attachment.getContentType().equals("image/gif")) {
					totalAttachment--;
				}
			}
		}
		return totalAttachment;
	}

	// Requirement 3
	// Get attachments name
	private List<String> getAttachmentsName(EmailMessage emailMessage) throws ServiceLocalException {
		List<String> attachmentsName = new ArrayList<>();
		
		// To check and read attachment name
		if (emailMessage.getHasAttachments()) {
			AttachmentCollection attachmentCollection = emailMessage.getAttachments();
			for (Attachment attachment : attachmentCollection.getItems()) {
				System.out.println("attachment.getContentType() ==>>" + attachment.getContentType());
				if (!attachment.getContentType().equals("image/png")
						&& !attachment.getContentType().equals("image/jpeg")
						&& !attachment.getContentType().equals("image/jpg")
						&& !attachment.getContentType().equals("image/gif")) {
					attachmentsName.add(attachment.getName());
				}
			}
		}
		return attachmentsName;
	}

	public static void main(String[] args) {
		String storeEmailPath = "E:\\client\\files\\email\\";

		ExchangeEmailTestAttachmentCount exchangeEmailTest = new ExchangeEmailTestAttachmentCount();

		List<String> emailList = new ArrayList<String>();
		emailList.add("INTERNAL_MAIL_ID");

		exchangeEmailTest.readEmails(emailList, storeEmailPath);
	}
}
