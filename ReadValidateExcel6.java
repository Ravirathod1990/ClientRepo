import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadValidateExcel6 {

	private List<TdoReportRow> listOfTdoDataFromReport = new ArrayList<TdoReportRow>();
	private List<CadacReportRow> listOfCadacDataFromReport = new ArrayList<CadacReportRow>();

	public String readAndValidateExcelData(String tdoFilepath, String tdoSheetName, String folderPathToFetchFiles,
			String targetFilePath) throws Exception {

		if (!new File(tdoFilepath).exists()) {
			return "TDO file not exist in directory";
		}
		readTdoReportExcel(tdoFilepath, tdoSheetName);
		readCadacReportExcel(folderPathToFetchFiles);
		return validateData(targetFilePath, tdoSheetName);
	}

	public void readTdoReportExcel(String filePath, String tdoSheetName) throws Exception {
		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));

		Sheet sheet = workbook.getSheet(tdoSheetName);

		Map<String, Integer> map = new HashMap<String, Integer>();
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		for (int x = 1; x < sheet.getPhysicalNumberOfRows(); x++) {
			TdoReportRow reportRow = new TdoReportRow();
			Row dataRow = sheet.getRow(x);

			int idxForVersion = map.get("Version");
			int idxForAgencyId = map.get("AgencyID");
			int idxForState = map.get("State");
			int idxForAgencyName = map.get("AgencyName");
			int idxSheetName = map.get("SheetName");

			Cell versionCell = dataRow.getCell(idxForVersion);
			Cell agencyIdCell = dataRow.getCell(idxForAgencyId);
			Cell stateCell = dataRow.getCell(idxForState);
			Cell agencyNameCell = dataRow.getCell(idxForAgencyName);
			Cell sheetNameCell = dataRow.getCell(idxSheetName);

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
			if (sheetNameCell != null) {
				reportRow.setSheetName(sheetNameCell.getStringCellValue());
			}

			listOfTdoDataFromReport.add(reportRow);
		}
	}

	public void readCadacReportExcel(String folderPathToFetchFiles) throws Exception {
		File dir = new File(folderPathToFetchFiles);
		for (TdoReportRow reportRow : listOfTdoDataFromReport) {
			if (reportRow.getSheetName() == null || reportRow.getSheetName().equals(StringUtils.EMPTY)) {
				return;
			}

			File[] files = dir.listFiles((d, name) -> name.startsWith(reportRow.getState()));
			if (files.length == 0) {
				return;
			}
			Workbook workbook = WorkbookFactory.create(new FileInputStream(files[0]));
			Sheet sheet = workbook.getSheet(reportRow.getSheetName());

			Map<String, Integer> map = new HashMap<String, Integer>();
			Row row = sheet.getRow(0);

			int minColIx = row.getFirstCellNum();
			int maxColIx = row.getLastCellNum();

			for (int colIx = minColIx; colIx < maxColIx; colIx++) {
				Cell cell = row.getCell(colIx);
				map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
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
					cadacReportRow.setFileCount(checkCellTypeAndReturn(fileCountCell));
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
					cadacReportRow.setRecordLength(checkCellTypeAndReturn(recordLengthCell));
				}
				if (recordCountCell != null) {
					cadacReportRow.setRecordCount(checkCellTypeAndReturn(recordCountCell));
				}
				if (instCell != null) {
					cadacReportRow.setSpecInst(instCell.getStringCellValue());
				}
				cadacReportRow.setState(reportRow.getState());
				listOfCadacDataFromReport.add(cadacReportRow);
			}
		}
	}

	public String checkCellTypeAndReturn(Cell cell) {
		String cellVal = StringUtils.EMPTY;
		switch (cell.getCellType()) {
		case NUMERIC:
			cellVal = String.valueOf(((int) cell.getNumericCellValue()));
			break;
		case STRING:
			cellVal = cell.getStringCellValue();
			break;
		default:
			break;
		}
		return cellVal;
	}

	public String validateData(String targetFilePath, String tdoSheetName) throws IOException {
		Map<Integer, Object[]> tdoMap = new TreeMap<Integer, Object[]>();
		int key = 1;
		tdoMap.put(key, new Object[] { "Version", "AgencyID", "State", "AgencyName", "Processs Type", "File Count",
				"File Type", "File Format", "Record Length", "Record Count" });

		StringBuffer result = new StringBuffer();
		for (TdoReportRow reportRow : listOfTdoDataFromReport) {
			boolean isFound = false;
			if (reportRow.getAgencyId() != null) {
				for (CadacReportRow cadacReportRow : listOfCadacDataFromReport) {
					if ((cadacReportRow.getTaxAuthorityId() != null
							&& reportRow.getAgencyId().equals(cadacReportRow.getTaxAuthorityId()))
							&& (reportRow.getState() != null && !reportRow.getState().equals(StringUtils.EMPTY)
									&& reportRow.getState().equals(cadacReportRow.getState()))) {
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
									reportRow.getAgencyName(), "no file /sheet exist", StringUtils.EMPTY,
									StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY });

					result.append(reportRow.getAgencyId() + " not found in file");
					result.append(System.lineSeparator());
				}
			}
		}
		writeDataToExcel(targetFilePath, tdoSheetName, tdoMap);
		return result.toString();
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

		String tdoFilepath = "E:\\client\\files\\requriment\\TDO2020.xlsx";
		String folderPathToFetchFiles = "E:\\client\\files\\requriment\\";
		String tdoSheetName = "BC3-YESFORMATV";
		String targetFilePath;

		ReadValidateExcel6 readValidateExcel = new ReadValidateExcel6();
		try {
			FileSystem system = FileSystems.getDefault();
			Path original = system.getPath(tdoFilepath);

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

			System.out.println(readValidateExcel.readAndValidateExcelData(tdoFilepath, tdoSheetName,
					folderPathToFetchFiles, targetFilePath));

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public class TdoReportRow {
		private String version;
		private String agencyId;
		private String state;
		private String agencyName;
		private String sheetName;

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

		public String getSheetName() {
			return sheetName;
		}

		public void setSheetName(String sheetName) {
			this.sheetName = sheetName;
		}
	}

	public class CadacReportRow {
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
		private String state;

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

		public String getState() {
			return state;
		}

		public void setState(String state) {
			this.state = state;
		}
	}
}
