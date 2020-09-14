import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FindNumber {

	private static final Pattern p = Pattern.compile("(?<!\\d)\\d{9}(?!\\d)");

	public static void main(String[] args) {
		FindNumber fn = new FindNumber();

		String str = "TAF 100909675 any string";
		String number = fn.extractNumberFromString(str);
		if (number != null) {
			System.out.println("number present in string ==>> " + number);
		} else {
			System.out.println("number not present in string");
		}

		boolean isExist = fn.checkTafTnl(str);
		System.out.println("TAF or TNL is exist ==>> " + isExist);
	}

	public String extractNumberFromString(String str) {
		String number = null;
		Matcher m = p.matcher(str);
		if (m.find()) {
			number = m.group(); // retrieve the matched substring
		}
		return number;
	}

	public boolean checkTafTnl(String str) {
		if (str.toLowerCase().contains("taf") || str.toLowerCase().contains("tnl")) {
			return true;
		}
		return false;
	}
}
