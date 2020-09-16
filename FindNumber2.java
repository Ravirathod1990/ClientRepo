import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FindNumber2 {

	private static final Pattern p = Pattern.compile("(?<!\\d)\\d{9}(?!\\d)");

	public static void main(String[] args) {
		FindNumber2 fn = new FindNumber2();

		// Requirement 1
		String str = "TAF 100909612 any 100909334 string";
		List<String> numberList1 = fn.extractNumbersFromString(str);
		if (!numberList1.isEmpty()) {
			System.out.println("number present in string ==>> ");
			for (String number : numberList1) {
				System.out.println(number);
			}
		} else {
			System.out.println("number not present in string");
		}

		// Requirement 4
		String str4 = "TAF 100909675 any 100909346 string 100909322 string any";
		List<String> numberList4 = fn.extractNumbersFromString(str4);
		if (!numberList4.isEmpty()) {
			String numberStr = "";
			for (String number : numberList4) {
				if (!numberStr.isEmpty()) {
					numberStr = numberStr+"+"+number;
				} else {
					numberStr = number;
				}
			}
			fn.renameSingleFile("E:\\client\\files\\new\\files\\newfolder-singlefile", numberStr, "PROD", "Sun Sep 13 10:48:37 CDT 2020");
		}

		// Requirement 5
		String str5 = "TAF 100909675 any string any";
		List<String> numberList5 = fn.extractNumbersFromString(str5);
		if (!numberList5.isEmpty()) {
			if (numberList5.size() == 1) {
				String agencyId = numberList5.get(0);
				fn.renameMultipleFiles("E:\\client\\files\\new\\files\\newfolder-multipleFile", agencyId, "PROD", "Sun Sep 13 10:48:37 CDT 2020");
			}
		}
	}

	public List<String> extractNumbersFromString(String str) {
		List<String> numbers = new ArrayList<>();
		Matcher m = p.matcher(str);
		while (m.find()) {
			numbers.add(m.group()); // retrieve the matched substring
		}
		return numbers;
	}

	public String convertDate(String dateStr) {
		String formattedDate = null;
		DateFormat originalFormat = new SimpleDateFormat("EEE MMM dd HH:mm:ss Z yyyy"); // Sun Sep 13 10:48:37 CDT 2020
		DateFormat targetFormat = new SimpleDateFormat("yyMMdd");
		Date date;
		try {
			date = originalFormat.parse(dateStr);
			formattedDate = targetFormat.format(date);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return formattedDate;
	}

	public void renameSingleFile(String folderPath, String agencyId, String prodType, String dateStr) {
		try {
			File folder = new File(folderPath);
			String fileName = folder.listFiles()[0].getName();
			String folderFileExtension = fileName.substring(fileName.lastIndexOf("."), fileName.length());
			String folderFileName = fileName.substring(0, fileName.lastIndexOf("."));
			String formattedDate = convertDate(dateStr);
			String actualFileName = agencyId + "_" + prodType + "_" + formattedDate;

			if (folderFileName.equals(actualFileName)) {
				System.out.println("File with name already exist in folder...");
			} else {
				Path filePath = Paths.get(folderPath+"\\"+fileName);
				Files.move(filePath, filePath.resolveSibling(actualFileName + folderFileExtension), StandardCopyOption.REPLACE_EXISTING);
				System.out.println("File name renamed successful...");
			}
		} catch(Exception e) {
			System.out.println(e);
		}
	}

	public void renameMultipleFiles(String folderPath, String agencyId, String prodType, String dateStr) {
		try {
			File folder = new File(folderPath);
			int i = 0;
			for (File file : folder.listFiles()) {
				String fileName = file.getName();
				String folderFileExtension = fileName.substring(fileName.lastIndexOf("."), fileName.length());
				String folderFileName = fileName.substring(0, fileName.lastIndexOf("."));
				String formattedDate = convertDate(dateStr);
				String actualFileName = agencyId + "-" + i + "_" + prodType + "_" + formattedDate;
				
				if (folderFileName.equals(actualFileName)) {
					System.out.println("File with name already exist in folder...");
				} else {
					Path filePath = Paths.get(folderPath+"\\"+fileName);
					Files.move(filePath, filePath.resolveSibling(actualFileName + folderFileExtension), StandardCopyOption.REPLACE_EXISTING);
					System.out.println("File name renamed successful...");
				}
				i++;
			}
		} catch(Exception e) {
			System.out.println(e);
		}
	}
}
