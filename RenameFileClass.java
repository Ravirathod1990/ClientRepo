import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class RenameFileClass {

	public static void main(String[] args) {
		RenameFileClass renameFileClass = new RenameFileClass();

		renameFileClass.renameFile("E:\\client\\files\\new\\files\\newfolder", "009089773", "PROD", "Sun Sep 13 10:48:37 CDT 2020");
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

	public void renameFile(String folderPath, String agencyId, String prodType, String dateStr) {
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
}
