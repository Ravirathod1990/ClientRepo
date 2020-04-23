import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.UUID;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadValidateExcel9 {

	private Map<String, CadacReportRow> listOfCadacDataFromReport = new HashMap<String, CadacReportRow>();
	private Map<String, TarTnlReportRow> listOfTarTnlFromReport = new HashMap<String, TarTnlReportRow>();
	private List<String> tarTnlColList = new ArrayList<>();
	private List<CadacReportRow> listOfTdoCopyData = new ArrayList<>();

	public void readAndValidateExcelData(String tdoFilepath, String cadacSheetName, String tarNtlSheetName,
			String targetFolder) throws Exception {

		String cadacFilepath = null;
		String tarNtlFilepath = null;

		File folder = new File(targetFolder);
		String[] allFiles = folder.list();
		for (String filename : allFiles) {
			if (filename.contains("DAC")) {
				cadacFilepath = targetFolder + filename;
			} else {
				tarNtlFilepath = targetFolder + filename;
			}
		}

		if (!new File(tdoFilepath).exists()) {
			System.out.println("TDO file not exist in directory");
			return;
		}
		if (!new File(cadacFilepath).exists()) {
			System.out.println("CA-DAC file not exist in directory");
			return;
		}
		readCadacReportExcel(cadacFilepath, cadacSheetName);
		readTarTnlReportExcel(tarNtlFilepath, tarNtlSheetName);
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
			map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			CadacReportRow cadacReportRow = new CadacReportRow();
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
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
					cadacReportRow.setTaxAuthorityId(checkCellTypeAndReturn(taxAuthIdCell));
				}
				if (vendorCell != null) {
					cadacReportRow.setVendor(checkCellTypeAndReturn(vendorCell));
				}
				if (processTypeCell != null) {
					cadacReportRow.setProcessType(checkCellTypeAndReturn(processTypeCell));
				}
				if (fileCountCell != null) {
					cadacReportRow.setFileCount(checkCellTypeAndReturn(fileCountCell));
				}
				if (fileTypeCell != null) {
					cadacReportRow.setFileType(checkCellTypeAndReturn(fileTypeCell));
				}
				if (convrCell != null) {
					cadacReportRow.setConversion(checkCellTypeAndReturn(convrCell));
				}
				if (fileFormatCell != null) {
					cadacReportRow.setFileFormat(checkCellTypeAndReturn(fileFormatCell));
				}
				if (recordLengthCell != null) {
					cadacReportRow.setRecordLength(checkCellTypeAndReturn(recordLengthCell));
				}
				if (recordCountCell != null) {
					cadacReportRow.setRecordCount(checkCellTypeAndReturn(recordCountCell));
				}
				if (instCell != null) {
					cadacReportRow.setSpecInst(checkCellTypeAndReturn(instCell));
				}
				if (cadacReportRow.getTaxAuthorityId() != null) {
					listOfCadacDataFromReport.put(cadacReportRow.getTaxAuthorityId(), cadacReportRow);
				}
			} else {
				physicalNumberOfRows++;
			}
		}
	}

	public void readTarTnlReportExcel(String filePath, String tarTnlSheetName) throws Exception {
		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));

		Sheet sheet = workbook.getSheet(tarTnlSheetName);

		Map<String, Integer> map = new HashMap<String, Integer>();
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		String agencyIdStr = null;
		String trspStr = null;
		String skipUpdatesStr = null;
		String paymentUpdatesStr = null;
		String othSepcialInstStr = null;

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			if (cell.getStringCellValue().trim().toLowerCase().contains("agency")) {
				agencyIdStr = cell.getStringCellValue();
			} else if (cell.getStringCellValue().trim().toLowerCase().contains("trsp recl")) {
				trspStr = cell.getStringCellValue();
			} else if (cell.getStringCellValue().trim().toLowerCase().contains("skip")) {
				skipUpdatesStr = cell.getStringCellValue();
			} else if (cell.getStringCellValue().trim().toLowerCase().contains("payment")) {
				paymentUpdatesStr = cell.getStringCellValue();
			} else if (cell.getStringCellValue().trim().toLowerCase().contains("special instructions")) {
				othSepcialInstStr = cell.getStringCellValue();
			}
			map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		tarTnlColList.add(trspStr);
		tarTnlColList.add(skipUpdatesStr);
		tarTnlColList.add(paymentUpdatesStr);
		tarTnlColList.add(othSepcialInstStr);
		tarTnlColList.add("Liability");

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			TarTnlReportRow tarTnlReportRow = new TarTnlReportRow();
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForAgencyId = map.get(agencyIdStr);
				int idxForTrsp = map.get(trspStr);
				int idxForSkipUpdates = map.get(skipUpdatesStr);
				int idxForPaymentUpdates = map.get(paymentUpdatesStr);
				int idxForSpecInst = map.get(othSepcialInstStr);

				Cell agencyIdCell = dataRow.getCell(idxForAgencyId);
				Cell trspCell = dataRow.getCell(idxForTrsp);
				Cell skipUpdatesCell = dataRow.getCell(idxForSkipUpdates);
				Cell paymentUpdatesCell = dataRow.getCell(idxForPaymentUpdates);
				Cell specInstCell = dataRow.getCell(idxForSpecInst);

				if (agencyIdCell != null) {
					String agencyIdVal = checkCellTypeAndReturn(agencyIdCell);
					if (agencyIdVal.toLowerCase().endsWith("x")) {
						agencyIdVal = agencyIdVal + "_" + UUID.randomUUID();
					}
					tarTnlReportRow.setAgencyId(agencyIdVal);
					if (trspCell != null) {
						tarTnlReportRow.setTrsp(checkCellTypeAndReturn(trspCell));
					}
					if (skipUpdatesCell != null) {
						tarTnlReportRow.setSkipUpdates(checkCellTypeAndReturn(skipUpdatesCell));
					}
					if (paymentUpdatesCell != null) {
						tarTnlReportRow.setPaymentUpdates(checkCellTypeAndReturn(paymentUpdatesCell));
					}
					if (specInstCell != null) {
						tarTnlReportRow.setOthSepcialInst(checkCellTypeAndReturn(specInstCell));
					}

					String othSpecialInst = tarTnlReportRow.getOthSepcialInst();
					if (tarTnlReportRow.getSkipUpdates() == null || tarTnlReportRow.getSkipUpdates().trim().isEmpty()) {
						if (othSpecialInst != null) {
							if (othSpecialInst.toLowerCase().contains("skip")) {
								tarTnlReportRow.setSkipUpdates(othSpecialInst);
							}
						}
					}

					if (othSpecialInst != null && othSpecialInst.toLowerCase().contains("high liability")) {
						int value = Integer.parseInt(othSpecialInst.replaceAll("[^0-9]", ""));
						tarTnlReportRow.setLiability(String.valueOf(value));
					}

					listOfTarTnlFromReport.put(tarTnlReportRow.getAgencyId(), tarTnlReportRow);
				}
			} else {
				physicalNumberOfRows++;
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

	public Map<String, Object[]> validateTdoData() throws IOException {
		Map<String, Object[]> tdoMap = new HashMap<String, Object[]>();

		for (Map.Entry<String, CadacReportRow> entry : listOfCadacDataFromReport.entrySet()) {
			CadacReportRow cadacReportRow = entry.getValue();
			if (cadacReportRow.getProcessType() != null && (cadacReportRow.getProcessType().equals("TAF")
					|| cadacReportRow.getProcessType().equals("TNL"))) {
				boolean isFileTypeExist = false;
				if (cadacReportRow.getFileType() != null && (cadacReportRow.getFileType().equals("txt")
						|| cadacReportRow.getFileType().equals("csv") || cadacReportRow.getFileType().equals("doc"))) {
					isFileTypeExist = true;

					tdoMap.put(entry.getKey(),
							new Object[] { cadacReportRow.getProcessType(), cadacReportRow.getFileCount(),
									cadacReportRow.getFileType(), cadacReportRow.getFileFormat(),
									cadacReportRow.getRecordLength(), cadacReportRow.getRecordCount() });
				} else if (cadacReportRow.getFileFormat() != null
						&& (cadacReportRow.getFileFormat().equals("txt") || cadacReportRow.getFileFormat().equals("csv")
								|| cadacReportRow.getFileFormat().equals("doc"))) {
					isFileTypeExist = true;

					tdoMap.put(entry.getKey(),
							new Object[] { cadacReportRow.getProcessType(), cadacReportRow.getFileCount(),
									cadacReportRow.getFileType(), cadacReportRow.getFileFormat(),
									cadacReportRow.getRecordLength(), cadacReportRow.getRecordCount() });
				}
				if (!isFileTypeExist) {
					tdoMap.put(entry.getKey(), new Object[] { cadacReportRow.getProcessType(), StringUtils.EMPTY,
							"not valid", "not valid", StringUtils.EMPTY, StringUtils.EMPTY });
				}
			} else {
				tdoMap.put(entry.getKey(), new Object[] { "not valid", StringUtils.EMPTY, StringUtils.EMPTY,
						StringUtils.EMPTY, StringUtils.EMPTY, StringUtils.EMPTY });
			}
		}
		return tdoMap;
	}

	public Map<String, Object[]> validateTarTnlData() throws IOException {
		Map<String, Object[]> tarTnlMap = new HashMap<String, Object[]>();

		for (Map.Entry<String, TarTnlReportRow> entry : listOfTarTnlFromReport.entrySet()) {
			TarTnlReportRow tarTnlReportRow = entry.getValue();
			tarTnlMap.put(entry.getKey(),
					new Object[] { tarTnlReportRow.getTrsp(), tarTnlReportRow.getSkipUpdates(),
							tarTnlReportRow.getPaymentUpdates(), tarTnlReportRow.getOthSepcialInst(),
							tarTnlReportRow.getLiability() });
		}
		return tarTnlMap;
	}

	public void copyFileSheet(String tdoFileName, String tdoFileNameCopy, String sheetName) throws IOException {
		Map<String, Object[]> tdoMap = validateTdoData();
		Map<String, Object[]> tarTnlMap = validateTarTnlData();

		BufferedInputStream bis = new BufferedInputStream(new FileInputStream(tdoFileName));
		XSSFWorkbook workbook = new XSSFWorkbook(bis);
		XSSFWorkbook myWorkBook = new XSSFWorkbook();
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		XSSFSheet mySheet = null;
		XSSFRow myRow = null;
		XSSFCell myCell = null;
		int fCell = 0;
		int lCell = 0;
		int headerlCell = 0;
		int fRow = 0;
		int lRow = 0;
		int idxForAgencyId = 0;
		int origCellCounterBeforeTnl = 0;

		sheet = workbook.getSheet(sheetName);
		if (sheet != null) {

			mySheet = myWorkBook.createSheet(sheet.getSheetName());
			fRow = sheet.getFirstRowNum();
			lRow = sheet.getLastRowNum();
			for (int iRow = fRow; iRow <= lRow; iRow++) {
				row = sheet.getRow(iRow);
				myRow = mySheet.createRow(iRow);
				if (row != null) {
					fCell = row.getFirstCellNum();
					lCell = row.getLastCellNum();
					if (iRow == 0) {
						headerlCell = lCell;
					}
					if (lCell != headerlCell) {
						lCell = headerlCell;
					}
					for (int iCell = fCell; iCell < lCell; iCell++) {
						cell = row.getCell(iCell);
						myCell = myRow.createCell(iCell);
						if (cell != null) {
							myCell.setCellType(cell.getCellType());
							myCell.setCellValue(checkCellTypeAndReturn(cell));

							if (iRow == 0 && checkCellTypeAndReturn(cell).equalsIgnoreCase("AgencyID")) {
								idxForAgencyId = cell.getColumnIndex();
							}
						}
					}
					if (iRow == 0) {
						fCell = lCell;

						List<String> list = new ArrayList<>();
						list.add("Processs Type");
						list.add("File Count");
						list.add("File Type");
						list.add("File Format");
						list.add("Record Length");
						list.add("Record Count");
						list.add("File Count DB");
						list.add("File Type DB");
						list.add("Record Length DB");
						list.add("Record Count DB");

						origCellCounterBeforeTnl = lCell + list.size();
						lCell = lCell + 15;

						for (String tarTnlColum : tarTnlColList) {
							list.add(tarTnlColum);
						}

						int counter = 0;
						for (int iCell = fCell; iCell < lCell; iCell++) {
							myCell = myRow.createCell(iCell);
							myCell.setCellType(CellType.STRING);
							myCell.setCellValue(list.get(counter));
							counter++;
						}
					} else {
						fCell = lCell;
						lCell = lCell + 6;

						Cell agencyIdCell = row.getCell(idxForAgencyId);
						String agencyValue = checkCellTypeAndReturn(agencyIdCell);
						Object[] tdoObj = tdoMap.get(agencyValue);
						if (tdoObj != null) {
							int counter = 0;
							for (int iCell = fCell; iCell < lCell; iCell++) {
								myCell = myRow.createCell(iCell);
								myCell.setCellType(CellType.STRING);
								if (tdoObj[counter] != null) {
									myCell.setCellValue(tdoObj[counter].toString());
								}
								counter++;
							}
						} else {
							String str = "not found in file";
							myCell = myRow.createCell(fCell);
							myCell.setCellType(CellType.STRING);
							myCell.setCellValue(str);
						}

						fCell = lCell;
						lCell = lCell + 4;
						List<String> dbRecordCountList = getDbRecordCount();
						int counter = 0;
						for (int iCell = fCell; iCell < lCell; iCell++) {
							myCell = myRow.createCell(iCell);
							myCell.setCellType(CellType.STRING);
							if (dbRecordCountList.get(counter) != null) {
								myCell.setCellValue(dbRecordCountList.get(counter));
							}
							counter++;
						}

						fCell = lCell;
						lCell = lCell + 5;
						Object[] tarTnlObj = tarTnlMap.get(agencyValue);
						if (tarTnlObj != null) {
							counter = 0;
							for (int iCell = fCell; iCell < lCell; iCell++) {
								myCell = myRow.createCell(iCell);
								myCell.setCellType(CellType.STRING);
								if (tarTnlObj[counter] != null) {
									myCell.setCellValue(tarTnlObj[counter].toString());
								}
								counter++;
							}
						}
					}
				}
			}

			int counter = lRow + 1;
			for (Map.Entry<String, TarTnlReportRow> entry : listOfTarTnlFromReport.entrySet()) {
				if (entry.getKey().toLowerCase().contains("x")) {
					TarTnlReportRow tarTnlReportRow = entry.getValue();
					int cellCounterBeforeTnl = origCellCounterBeforeTnl;
					myRow = mySheet.createRow(counter);
					myCell = myRow.createCell(idxForAgencyId);
					myCell.setCellType(CellType.STRING);
					myCell.setCellValue(tarTnlReportRow.getAgencyId().split("_")[0]);

					myCell = myRow.createCell(cellCounterBeforeTnl);
					myCell.setCellType(CellType.STRING);
					myCell.setCellValue(tarTnlReportRow.getTrsp());

					cellCounterBeforeTnl++;
					myCell = myRow.createCell(cellCounterBeforeTnl);
					myCell.setCellType(CellType.STRING);
					myCell.setCellValue(tarTnlReportRow.getSkipUpdates());

					cellCounterBeforeTnl++;
					myCell = myRow.createCell(cellCounterBeforeTnl);
					myCell.setCellType(CellType.STRING);
					myCell.setCellValue(tarTnlReportRow.getPaymentUpdates());

					cellCounterBeforeTnl++;
					myCell = myRow.createCell(cellCounterBeforeTnl);
					myCell.setCellType(CellType.STRING);
					myCell.setCellValue(tarTnlReportRow.getOthSepcialInst());

					cellCounterBeforeTnl++;
					myCell = myRow.createCell(cellCounterBeforeTnl);
					myCell.setCellType(CellType.STRING);
					myCell.setCellValue(tarTnlReportRow.getLiability());

					counter++;
				}
			}
		}
		FileOutputStream fileOut = new FileOutputStream(new File(tdoFileNameCopy));
		myWorkBook.write(fileOut);
		bis.close();
		workbook.close();
		myWorkBook.close();
		fileOut.close();
	}

	public void readTdoCopyFileAndInsertDataIntoDB(String filePath, String tdoSheetName) throws Exception {
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

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			CadacReportRow tdoCopyFileRow = new CadacReportRow();
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForTaxAuthId = map.get("AgencyID");
				int idxForFileCount = map.get("File Count");
				int idxForFileType = map.get("File Type");
				int idxForRecordLength = map.get("Record Length");
				int idxForRecordCount = map.get("Record Count");

				Cell taxAuthIdCell = dataRow.getCell(idxForTaxAuthId);
				Cell fileCountCell = dataRow.getCell(idxForFileCount);
				Cell fileTypeCell = dataRow.getCell(idxForFileType);
				Cell recordLengthCell = dataRow.getCell(idxForRecordLength);
				Cell recordCountCell = dataRow.getCell(idxForRecordCount);

				if (taxAuthIdCell != null) {
					tdoCopyFileRow.setTaxAuthorityId(checkCellTypeAndReturn(taxAuthIdCell));
				}
				if (fileCountCell != null) {
					tdoCopyFileRow.setFileCount(checkCellTypeAndReturn(fileCountCell));
				}
				if (fileTypeCell != null) {
					tdoCopyFileRow.setFileType(checkCellTypeAndReturn(fileTypeCell));
				}
				if (recordLengthCell != null) {
					tdoCopyFileRow.setRecordLength(checkCellTypeAndReturn(recordLengthCell));
				}
				if (recordCountCell != null) {
					tdoCopyFileRow.setRecordCount(checkCellTypeAndReturn(recordCountCell));
				}
				if (tdoCopyFileRow.getTaxAuthorityId() != null) {
					listOfTdoCopyData.add(tdoCopyFileRow);
				}
			} else {
				physicalNumberOfRows++;
			}
		}

		for (CadacReportRow tdoCopyFileRow : listOfTdoCopyData) {
			if (tdoCopyFileRow.getTaxAuthorityId() != null
					&& !tdoCopyFileRow.getTaxAuthorityId().toLowerCase().endsWith("x")) {
				String agencyId = tdoCopyFileRow.getTaxAuthorityId();
				String fileType = tdoCopyFileRow.getFileType();
				String fileCount = tdoCopyFileRow.getFileCount();
				String recordCount = tdoCopyFileRow.getRecordCount();
				String recordLength = tdoCopyFileRow.getRecordLength();

				// It's for reference purpose, if not used then remove it
				System.out.println("Agency ID ==>>" + agencyId);
				System.out.println("File Type ==>>" + fileType);
				System.out.println("File Count ==>>" + fileCount);
				System.out.println("Record Count ==>>" + recordCount);
				System.out.println("Record Length ==>>" + recordLength);

				// Query to check into database
				Integer queryResult = 1;
				// Code to fetch data from the database

				if (queryResult > 0) {
					// Code to update into database
				} else {
					// Code to insert into database
				}
			}
		}
	}

	public List<String> getDbRecordCount() {
		Random rand = new Random();
		List<String> dbRecordCountList = new ArrayList<>();
		dbRecordCountList.add(String.valueOf(rand.nextInt(1000))); // File Count DB
		dbRecordCountList.add(String.valueOf(rand.nextInt(1000))); // File Type DB
		dbRecordCountList.add(String.valueOf(rand.nextInt(1000))); // Record Length DB
		dbRecordCountList.add(String.valueOf(rand.nextInt(1000))); // Record Count DB
		return dbRecordCountList;
	}

	public static void main(String[] args) {

		String tdoFilepath = "E:\\client\\files\\newrequriemnt\\TDO2020.xlsx";
		String targetFolder = "E:\\client\\files\\newrequriemnt\\sample\\";
		String tdoSheetName = "BC3-YESFORMATV";
		String cadacSheetName = "STATE MATRIX";
		String tarNtlSheetName = "CT";
		String targetFilePath;

		ReadValidateExcel9 readValidateExcel = new ReadValidateExcel9();
		try {
			FileSystem system = FileSystems.getDefault();
			Path original = system.getPath(tdoFilepath);

			String origFileName = original.getFileName().toString();
			String fileName = origFileName.substring(0, origFileName.lastIndexOf("."));
			String fileExt = origFileName.substring(origFileName.lastIndexOf("."), origFileName.length());
			Path targetFile = system.getPath(original.getParent() + "\\" + fileName + "_copy" + fileExt);
			targetFilePath = targetFile.toString();

			try {
				readValidateExcel.readAndValidateExcelData(tdoFilepath, cadacSheetName, tarNtlSheetName, targetFolder);
				File file = new File(targetFilePath);
				file.createNewFile();
				readValidateExcel.copyFileSheet(tdoFilepath, targetFilePath, tdoSheetName);

				// Read data from TDO-COPY file and insert/update in database
				readValidateExcel.readTdoCopyFileAndInsertDataIntoDB(targetFilePath, tdoSheetName);
			} catch (IOException ex) {
				System.out.println(ex.getMessage());
			}

		} catch (Exception e) {
			e.printStackTrace();
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

	public class TarTnlReportRow {
		private String agencyId;
		private String trsp;
		private String skipUpdates;
		private String paymentUpdates;
		private String othSepcialInst;
		private String liability;

		public String getAgencyId() {
			return agencyId;
		}

		public void setAgencyId(String agencyId) {
			this.agencyId = agencyId;
		}

		public String getTrsp() {
			return trsp;
		}

		public void setTrsp(String trsp) {
			this.trsp = trsp;
		}

		public String getSkipUpdates() {
			return skipUpdates;
		}

		public void setSkipUpdates(String skipUpdates) {
			this.skipUpdates = skipUpdates;
		}

		public String getPaymentUpdates() {
			return paymentUpdates;
		}

		public void setPaymentUpdates(String paymentUpdates) {
			this.paymentUpdates = paymentUpdates;
		}

		public String getOthSepcialInst() {
			return othSepcialInst;
		}

		public void setOthSepcialInst(String othSepcialInst) {
			this.othSepcialInst = othSepcialInst;
		}

		public String getLiability() {
			return liability;
		}

		public void setLiability(String liability) {
			this.liability = liability;
		}
	}
}
