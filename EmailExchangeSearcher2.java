
import java.net.URI;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FolderView;

public class EmailExchangeSearcher2 {

	public static boolean sendEmail() {

		Boolean flag = false;
		try {

			ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
			ExchangeCredentials credentials = new WebCredentials("username", "password");
			service.setCredentials(credentials);
			service.setUrl(new URI("https://mail.hostname.com/EWS/Exchange.asmx"));

			findChildFolders(service);

			flag = true;
		} catch (Exception e) {
			e.printStackTrace();
		}

		return flag;

	}

	public static void findChildFolders(ExchangeService service) throws Exception {
		FindFoldersResults findResults = service.findFolders(WellKnownFolderName.Inbox,
				new FolderView(Integer.MAX_VALUE));

		for (Folder folder : findResults.getFolders()) {
			System.out.println("Count======" + folder.getChildFolderCount());
			System.out.println("Name=======" + folder.getDisplayName());
		}
	}

	public static void main(String[] args) {

		sendEmail();

	}

}
