import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

public class ReadvalidateData {

	public static void main(String[] args) throws IOException {

		List<String> list = Files.readAllLines(new File("E:\\client\\files\\sample.txt").toPath(),
				Charset.defaultCharset());

		String sampleStr = "";
		for (String line : list) {
			if (line.contains("SAMPLE DATA")) {
				sampleStr = line;
			}
		}

		String[] strArr = sampleStr.split(" ");
		List<String> strList = new ArrayList<>();
		for (String localStr : strArr) {
			String actStr = localStr.replaceAll("[^0-9]", "");
			if (actStr.trim().length() != 0) {
				strList.add(actStr);
			}
		}

		System.out.println(strList.get(strList.size() - 1));
	}

}
