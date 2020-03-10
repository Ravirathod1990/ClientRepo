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
		boolean isStart = false;
		boolean isSkip = false;

		try {
			list = Files.readAllLines(file.toPath(), Charset.defaultCharset());
			for (String line : list) {
				if (line.contains("1...5...10")) {
					if (!isStart) {
						isStart = true;
						isSkip = true;
					} else if (isStart) {
						isStart = false;
					}
				} else {
					isSkip = false;
				}
				if (isStart && !isSkip) {
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
