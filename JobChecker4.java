import java.io.File;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.util.List;

public class JobChecker4 {

	public static void main(String[] args) {
		File file = new File("E:\\client\\files\\REQUIREMENT1.txt");
		boolean exists = file.exists();

		if (!exists) {
			System.out.println("File not exist in the directory");
			return;
		}

		String finalResult = extract(file);
		System.out.println(finalResult);
	}

	private static String extract(File file) {

		List<String> list;
		StringBuffer extractedText = new StringBuffer();
		boolean isMainTaskStart = false;
		boolean isChildTaskStart = false;
		boolean isSkip = false;

		try {
			list = Files.readAllLines(file.toPath(), Charset.defaultCharset());
			for (String line : list) {
				if (line.contains("SAMPLE TAXIDS TABLE")) {
					isMainTaskStart = true;
				} else if (line.contains("1...5...10")) {
					if (!isChildTaskStart && isMainTaskStart) {
						isChildTaskStart = true;
						isSkip = true;
					} else if (isMainTaskStart && isChildTaskStart) {
						isMainTaskStart = false;
						isChildTaskStart = false;
					}
				} else {
					isSkip = false;
				}
				if (isMainTaskStart && isChildTaskStart && !isSkip) {
					extractedText.append(line);
					extractedText.append(System.lineSeparator());
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		return extractedText.toString();
	}
}
