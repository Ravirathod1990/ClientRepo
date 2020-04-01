import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class UnzipSearchDelete {

	private static final int BUFFER_SIZE = 4096;

	public void unzip(String zipFilePath, String destDirectory) throws IOException {

		File destDir = new File(destDirectory);
		if (!destDir.exists()) {
			destDir.mkdir();
		}
		ZipInputStream zipIn = new ZipInputStream(new FileInputStream(zipFilePath));
		ZipEntry entry = zipIn.getNextEntry();
		while (entry != null) {
			String filePath = destDirectory + File.separator + entry.getName();
			if (!entry.isDirectory()) {
				extractFile(zipIn, filePath);
			} else {
				File dir = new File(filePath);
				dir.mkdir();
			}
			zipIn.closeEntry();
			entry = zipIn.getNextEntry();
		}
		zipIn.close();
	}

	private void extractFile(ZipInputStream zipIn, String filePath) throws IOException {
		BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(filePath));
		byte[] bytesIn = new byte[BUFFER_SIZE];
		int read = 0;
		while ((read = zipIn.read(bytesIn)) != -1) {
			bos.write(bytesIn, 0, read);
		}
		bos.close();
	}

	public List<File> fileFinder(String dirName) {
		Path dir = Paths.get(dirName);

		List<File> files = new ArrayList<>();
		try (DirectoryStream<Path> stream = Files.newDirectoryStream(dir, "*DAC*.xlsx")) {
			for (Path entry : stream) {
				files.add(entry.toFile());
			}
			return files;
		} catch (IOException x) {
			throw new RuntimeException(String.format("error reading folder %s: %s", dir, x.getMessage()), x);
		}
	}

	public String createFolder(String zipFileName, String folderPath) {
		String orgFileName = zipFileName.replace(".zip", "");
		String zipFolderPath = new StringBuffer(folderPath).append(File.separator).append(orgFileName).toString();
		File dir = new File(zipFolderPath);
		dir.mkdir();
		return zipFolderPath;
	}

	public static void main(String[] args) {

		String folderPath = "E:\\client\\files\\requriment";
		String ext = ".zip";
		UnzipSearchDelete unzipSearchDelete = new UnzipSearchDelete();

		try {
			GenericExtFilter filter = unzipSearchDelete.new GenericExtFilter(ext);
			File dir = new File(folderPath);
			String[] list = dir.list(filter);

			if (list.length == 0) {
				System.out.println("no files end with : " + ext);
				return;
			}

			for (String file : list) {
				String zipFilePath = new StringBuffer(folderPath).append(File.separator).append(file).toString();
				String zipFolderPath = unzipSearchDelete.createFolder(file, folderPath);
				unzipSearchDelete.unzip(zipFilePath, zipFolderPath);
				List<File> fileList = unzipSearchDelete.fileFinder(zipFolderPath);
				File zipFolderPathFile = new File(zipFolderPath);
				for (File fileObj : zipFolderPathFile.listFiles()) {
					if (!fileList.contains(fileObj)) {
						fileObj.delete();
					}
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public class GenericExtFilter implements FilenameFilter {

		private String ext;

		public GenericExtFilter(String ext) {
			this.ext = ext;
		}

		public boolean accept(File dir, String name) {
			return (name.endsWith(ext));
		}
	}
}
