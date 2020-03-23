import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadValidateExcel2 {

	private List<TdoReportRow> listOfRdoDataFromReport = new ArrayList<TdoReportRow>();
	private List<CadacReportRow> listOfCadacDataFromReport = new ArrayList<CadacReportRow>();

	private static final int BUFFER_SIZE = 4096;

	public String readAndValidateExcelData(String tdoFilepath, String tdoSheetName, String cadacFilepath,
			String cadacSheetName, String targetFilePath) throws Exception {

		if (!new File(tdoFilepath).exists()) {
			return "TDO file not exist in directory";
		}
		if (!new File(cadacFilepath).exists()) {
			return "CA-DAC file not exist in directory";
		}
		readRdoReportExcel(tdoFilepath, tdoSheetName);
		readCadacReportExcel(cadacFilepath, cadacSheetName);
		return validateData(targetFilePath, tdoSheetName);
	}

	public void readRdoReportExcel(String filePath, String tdoSheetName) throws Exception {
		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));

		Sheet sheet = workbook.getSheet(tdoSheetName);

		Map<String, Integer> map = new HashMap<String, Integer>();
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			map.put(cell.getStringCellValue(), cell.getColumnIndex());
		}

		for (int x = 1; x < sheet.getPhysicalNumberOfRows(); x++) {
			TdoReportRow reportRow = new TdoReportRow();
			Row dataRow = sheet.getRow(x);

			int idxForVersion = map.get("Version");
			int idxForAgencyId = map.get("AgencyID");
			int idxForState = map.get("State");
			int idxForAgencyName = map.get("AgencyName");

			Cell versionCell = dataRow.getCell(idxForVersion);
			Cell agencyIdCell = dataRow.getCell(idxForAgencyId);
			Cell stateCell = dataRow.getCell(idxForState);
			Cell agencyNameCell = dataRow.getCell(idxForAgencyName);

			if (versionCell != null) {
				reportRow.setVersion(String.valueOf(((int) versionCell.getNumericCellValue())));
			}
			if (agencyIdCell != null) {
				reportRow.setAgencyId(agencyIdCell.getStringCellValue());
			}
			if (stateCell != null) {
				reportRow.setState(stateCell.getStringCellValue());
			}
			if (agencyNameCell != null) {
				reportRow.setAgencyName(agencyNameCell.getStringCellValue());
			}

			listOfRdoDataFromReport.add(reportRow);
		}
	}

	public void readCadacReportExcel(String filePath, String cadacSheetName) throws Exception {
		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));

		Sheet sheet = workbook.getSheet(cadacSheetName);

		Map<String, Integer> map = new HashMap<String, Integer>();
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			map.put(cell.getStringCellValue(), cell.getColumnIndex());
		}

		for (int x = 1; x < sheet.getPhysicalNumberOfRows(); x++) {
			CadacReportRow cadacReportRow = new CadacReportRow();
			Row dataRow = sheet.getRow(x);

			int idxForTaxAuthId = map.get("TAX AUTHORITY ID");
			int idxForVendor = map.get("VENDOR");
			int idxForProcessType = map.get("Processs Type");
			int idxForFileCount = map.get("File Count");
			int idxForFileType = map.get("File Type");
			int idxForConvr = map.get("Conversion");
			int idxForFileFormat = map.get("File Format");
			int idxForRecordLength = map.get("Record Length");
			int idxForRecordCount = map.get("Record Count");
			int idxForInst = map.get("Special Isntructions");

			Cell taxAuthIdCell = dataRow.getCell(idxForTaxAuthId);
			Cell vendorCell = dataRow.getCell(idxForVendor);
			Cell processTypeCell = dataRow.getCell(idxForProcessType);
			Cell fileCountCell = dataRow.getCell(idxForFileCount);
			Cell fileTypeCell = dataRow.getCell(idxForFileType);
			Cell convrCell = dataRow.getCell(idxForConvr);
			Cell fileFormatCell = dataRow.getCell(idxForFileFormat);
			Cell recordLengthCell = dataRow.getCell(idxForRecordLength);
			Cell recordCountCell = dataRow.getCell(idxForRecordCount);
			Cell instCell = dataRow.getCell(idxForInst);

			if (taxAuthIdCell != null) {
				cadacReportRow.setTaxAuthorityId(taxAuthIdCell.getStringCellValue());
			}
			if (vendorCell != null) {
				cadacReportRow.setVendor(vendorCell.getStringCellValue());
			}
			if (processTypeCell != null) {
				cadacReportRow.setProcessType(processTypeCell.getStringCellValue());
			}
			if (fileCountCell != null) {
				cadacReportRow.setFileCount(String.valueOf(((int) fileCountCell.getNumericCellValue())));
			}
			if (fileTypeCell != null) {
				cadacReportRow.setFileType(fileTypeCell.getStringCellValue());
			}
			if (convrCell != null) {
				cadacReportRow.setConversion(convrCell.getStringCellValue());
			}
			if (fileFormatCell != null) {
				cadacReportRow.setFileFormat(fileFormatCell.getStringCellValue());
			}
			if (recordLengthCell != null) {
				cadacReportRow.setRecordLength(String.valueOf(((int) recordLengthCell.getNumericCellValue())));
			}
			if (recordCountCell != null) {
				cadacReportRow.setRecordCount(String.valueOf(((int) recordCountCell.getNumericCellValue())));
			}
			if (instCell != null) {
				cadacReportRow.setSpecInst(instCell.getStringCellValue());
			}
			listOfCadacDataFromReport.add(cadacReportRow);
		}
	}

	public String validateData(String targetFilePath, String tdoSheetName) throws IOException {
		Map<Integer, Object[]> tdoMap = new TreeMap<Integer, Object[]>();
		int key = 1;
		tdoMap.put(key, new Object[] { "Version", "AgencyID", "State", "AgencyName", "Processs Type", "File Count",
				"File Type", "File Format", "Record Length", "Record Count" });

		StringBuffer result = new StringBuffer();
		for (TdoReportRow reportRow : listOfRdoDataFromReport) {
			boolean isFound = false;
			if (reportRow.getAgencyId() != null) {
				for (CadacReportRow cadacReportRow : listOfCadacDataFromReport) {
					if (cadacReportRow.getTaxAuthorityId() != null
							&& reportRow.getAgencyId().equals(cadacReportRow.getTaxAuthorityId())) {
						isFound = true;
						result.append(reportRow.getAgencyId());
						if (cadacReportRow.getProcessType() != null && (cadacReportRow.getProcessType().equals("TAF")
								|| cadacReportRow.getProcessType().equals("TNL"))) {
							boolean isFileTypeExist = false;
							if (cadacReportRow.getFileType() != null && (cadacReportRow.getFileType().equals("txt")
									|| cadacReportRow.getFileType().equals("csv")
									|| cadacReportRow.getFileType().equals("doc"))) {
								isFileTypeExist = true;
								result.append(" File Count:" + cadacReportRow.getFileCount());
								result.append(" File Type:" + cadacReportRow.getFileType());
								result.append(" Record Count:" + cadacReportRow.getRecordCount());
								result.append(" Record Length:" + cadacReportRow.getRecordLength());

								key = key + 1;
								tdoMap.put(key,
										new Object[] { reportRow.getVersion(), reportRow.getAgencyId(),
												reportRow.getState(), reportRow.getAgencyName(),
												cadacReportRow.getProcessType(), cadacReportRow.getFileCount(),
												cadacReportRow.getFileType(), cadacReportRow.getFileFormat(),
												cadacReportRow.getRecordLength(), cadacReportRow.getRecordCount() });
							} else if (cadacReportRow.getFileFormat() != null
									&& (cadacReportRow.getFileFormat().equals("txt")
											|| cadacReportRow.getFileFormat().equals("csv")
											|| cadacReportRow.getFileFormat().equals("doc"))) {
								isFileTypeExist = true;
								result.append(" File Count:" + cadacReportRow.getFileCount());
								result.append(" File Type:" + cadacReportRow.getFileFormat());
								result.append(" Record Count:" + cadacReportRow.getRecordCount());
								result.append(" Record Length:" + cadacReportRow.getRecordLength());

								key = key + 1;
								tdoMap.put(key,
										new Object[] { reportRow.getVersion(), reportRow.getAgencyId(),
												reportRow.getState(), reportRow.getAgencyName(),
												cadacReportRow.getProcessType(), cadacReportRow.getFileCount(),
												cadacReportRow.getFileType(), cadacReportRow.getFileFormat(),
												cadacReportRow.getRecordLength(), cadacReportRow.getRecordCount() });
							}
							if (!isFileTypeExist) {
								key = key + 1;
								tdoMap.put(key,
										new Object[] { reportRow.getVersion(), reportRow.getAgencyId(),
												reportRow.getState(), reportRow.getAgencyName(),
												cadacReportRow.getProcessType(), StringUtils.EMPTY, "not valid",
												"not valid", StringUtils.EMPTY, StringUtils.EMPTY });
								result.append(" File Type/File Format: not valid");
							}
						} else {
							key = key + 1;
							tdoMap.put(key, new Object[] { reportRow.getVersion(), reportRow.getAgencyId(),
									reportRow.getState(), reportRow.getAgencyName(), "not valid", StringUtils.EMPTY,
									StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY });
							result.append(" Processs Type: not valid");
						}
						result.append(System.lineSeparator());
						break;
					}
				}
				if (!isFound) {
					key = key + 1;
					tdoMap.put(key,
							new Object[] { reportRow.getVersion(), reportRow.getAgencyId(), reportRow.getState(),
									reportRow.getAgencyName(), reportRow.getAgencyId() + "not found in file",
									StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY,
									StringUtils.EMPTY });

					result.append(reportRow.getAgencyId() + " not found in file");
					result.append(System.lineSeparator());
				}
			}
		}
		writeDataToExcel(targetFilePath, tdoSheetName, tdoMap);
		return result.toString();
	}

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

	public File[] fileFinder(String dirName) {
		File dir = new File(dirName);

		return dir.listFiles(new FilenameFilter() {
			public boolean accept(File dir, String filename) {
				return filename.endsWith(".xlsx");
			}
		});
	}

	public void writeDataToExcel(String targetFilePath, String tdoSheetName, Map<Integer, Object[]> tdoMap)
			throws IOException {

		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(new FileInputStream(targetFilePath));

		XSSFSheet spreadsheet = workbook.getSheet(tdoSheetName);

		// Create row object
		XSSFRow row;

		// Iterate over data and write to sheet
		Set<Integer> keyid = tdoMap.keySet();
		int rowid = 0;

		for (Integer key : keyid) {
			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = tdoMap.get(key);
			int cellid = 0;

			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}

		// Write the workbook in file system
		FileOutputStream out = new FileOutputStream(new File(targetFilePath));

		workbook.write(out);
		out.close();
		System.out.println("Data written successfully");
	}

	public static void main(String[] args) {

		String zipFilePath = "E:\\client\\files\\extract\\files.zip";
		String fileExtractPath = "E:\\client\\files\\extract\\";
		String tdoFilePath = "E:\\client\\files\\TDO2020.xlsx";
		String tdoSheetName = "BC3-YESFORMATV";
		String cadacSheetName = "STATE MATRIX";
		String targetFilePath;

		ReadValidateExcel2 readValidateExcel = new ReadValidateExcel2();
		try {
			FileSystem system = FileSystems.getDefault();
			Path original = system.getPath(tdoFilePath);

			String origFileName = original.getFileName().toString();
			String fileName = origFileName.substring(0, origFileName.lastIndexOf("."));
			String fileExt = origFileName.substring(origFileName.lastIndexOf("."), origFileName.length());
			Path targetFile = system.getPath(original.getParent() + "\\" + fileName + "_copy" + fileExt);
			targetFilePath = targetFile.toString();

			try {
				Files.copy(original, targetFile, StandardCopyOption.REPLACE_EXISTING);
			} catch (IOException ex) {
				System.out.println(ex.getMessage());
			}

			readValidateExcel.unzip(zipFilePath, fileExtractPath);
			File[] fileArr = readValidateExcel.fileFinder(fileExtractPath);
			if (fileArr.length > 1) {
				System.out.println("Multiple file with .xlsx extension exist on the directory");
				return;
			} else {
				System.out.println(readValidateExcel.readAndValidateExcelData(tdoFilePath, tdoSheetName,
						fileArr[0].getPath(), cadacSheetName, targetFilePath));
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}

class TdoReportRow {
	private String version;
	private String agencyId;
	private String state;
	private String agencyName;

	public String getVersion() {
		return version;
	}

	public void setVersion(String version) {
		this.version = version;
	}

	public String getAgencyId() {
		return agencyId;
	}

	public void setAgencyId(String agencyId) {
		this.agencyId = agencyId;
	}

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}

	public String getAgencyName() {
		return agencyName;
	}

	public void setAgencyName(String agencyName) {
		this.agencyName = agencyName;
	}
}

class CadacReportRow {
	private String taxAuthorityId;
	private String vendor;
	private String processType;
	private String fileCount;
	private String fileType;
	private String conversion;
	private String fileFormat;
	private String recordLength;
	private String recordCount;
	private String specInst;

	public String getTaxAuthorityId() {
		return taxAuthorityId;
	}

	public void setTaxAuthorityId(String taxAuthorityId) {
		this.taxAuthorityId = taxAuthorityId;
	}

	public String getVendor() {
		return vendor;
	}

	public void setVendor(String vendor) {
		this.vendor = vendor;
	}

	public String getProcessType() {
		return processType;
	}

	public void setProcessType(String processType) {
		this.processType = processType;
	}

	public String getFileCount() {
		return fileCount;
	}

	public void setFileCount(String fileCount) {
		this.fileCount = fileCount;
	}

	public String getFileType() {
		return fileType;
	}

	public void setFileType(String fileType) {
		this.fileType = fileType;
	}

	public String getConversion() {
		return conversion;
	}

	public void setConversion(String conversion) {
		this.conversion = conversion;
	}

	public String getFileFormat() {
		return fileFormat;
	}

	public void setFileFormat(String fileFormat) {
		this.fileFormat = fileFormat;
	}

	public String getRecordLength() {
		return recordLength;
	}

	public void setRecordLength(String recordLength) {
		this.recordLength = recordLength;
	}

	public String getRecordCount() {
		return recordCount;
	}

	public void setRecordCount(String recordCount) {
		this.recordCount = recordCount;
	}

	public String getSpecInst() {
		return specInst;
	}

	public void setSpecInst(String specInst) {
		this.specInst = specInst;
	}
}