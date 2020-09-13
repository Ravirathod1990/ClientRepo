import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateConverter {

	public static void main(String[] args) {

		DateFormat originalFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss aaa yyyy"); // Sun Sep 13 10:48:37 AM 2020
		DateFormat targetFormat = new SimpleDateFormat("yyMMdd");
		Date date;
		try {
			date = originalFormat.parse("Sun Sep 13 10:48:37 AM 2020");
			String formattedDate = targetFormat.format(date);
			System.out.println("Date ==>>" + formattedDate);
		} catch (ParseException e) {
			e.printStackTrace();
		}
	}
}
