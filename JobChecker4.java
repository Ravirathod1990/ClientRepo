import java.io.File;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

public class JobChecker4 {

	public static void main(String[] args) {
		String result = searchFile("E:\\client\\files");
		System.out.println(result);

		if (result.equals("File not exist in the directory")) {
			return;
		}

		String finalResult = extract(result);
		System.out.println(finalResult);
	}

	private static String extract(String result) {

		List<String> list;
		StringBuffer extractedText = new StringBuffer();
		boolean isStart = false;
		boolean isSkip = false;

		try {
			list = Files.readAllLines(new File(result).toPath(), Charset.defaultCharset());
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

	public static String searchFile(String folderLocation) {
		final File folder = new File(folderLocation);
		List<String> fileList = new ArrayList<>();
		search(".*\\.txt", folder, fileList);
		if (fileList.size() == 0) {
			return "File not exist in the directory";
		} else {
			return fileList.get(0);
		}
	}

	public static void search(final String pattern, final File folder, List<String> fileList) {
		for (final File f : folder.listFiles()) {

			if (f.isDirectory()) {
				search(pattern, f, fileList);
			}

			if (f.isFile()) {
				if (f.getName().matches(pattern)) {
					fileList.add(f.getAbsolutePath());
				}
			}
		}
	}
}
