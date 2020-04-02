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

public class ReadValidateExcel3 {

	private Map<String, CadacReportRow> listOfCadacDataFromReport = new HashMap<String, CadacReportRow>();

	public void readAndValidateExcelData(String tdoFilepath, String cadacFilepath, String cadacSheetName)
			throws Exception {

		if (!new File(tdoFilepath).exists()) {
			System.out.println("TDO file not exist in directory");
			return;
		}
		if (!new File(cadacFilepath).exists()) {
			System.out.println("CA-DAC file not exist in directory");
			return;
		}
		readCadacReportExcel(cadacFilepath, cadacSheetName);
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
			if (cadacReportRow.getTaxAuthorityId() != null) {
				listOfCadacDataFromReport.put(cadacReportRow.getTaxAuthorityId(), cadacReportRow);
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

	public Map<String, Object[]> validateData() throws IOException {
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

	public void copyFileSheet(String tdoFileName, String tdoFileNameCopy, String sheetName) throws IOException {
		Map<String, Object[]> tdoMap = validateData();

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
		int fRow = 0;
		int lRow = 0;
		int idxForAgencyId = 0;

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
						lCell = lCell + 6;

						List<String> list = new ArrayList<>();
						list.add("Processs Type");
						list.add("File Count");
						list.add("File Type");
						list.add("File Format");
						list.add("Record Length");
						list.add("Record Count");

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
						Object[] obj = tdoMap.get(agencyValue);
						if (obj != null) {
							int counter = 0;
							for (int iCell = fCell; iCell < lCell; iCell++) {
								myCell = myRow.createCell(iCell);
								myCell.setCellType(CellType.STRING);
								if (obj[counter] != null) {
									myCell.setCellValue(obj[counter].toString());
								}
								counter++;
							}
						} else {
							String str = "not found in file";
							myCell = myRow.createCell(fCell);
							myCell.setCellType(CellType.STRING);
							myCell.setCellValue(str);
						}
					}
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

	public static void main(String[] args) {

		String tdoFilepath = "E:\\client\\files\\TDO2020.xlsx";
		String cadacFilepath = "E:\\client\\files\\CA-DAC.xlsx";
		String tdoSheetName = "BC3-YESFORMATV";
		String cadacSheetName = "STATE MATRIX";
		String targetFilePath;

		ReadValidateExcel3 readValidateExcel = new ReadValidateExcel3();
		try {
			FileSystem system = FileSystems.getDefault();
			Path original = system.getPath(tdoFilepath);

			String origFileName = original.getFileName().toString();
			String fileName = origFileName.substring(0, origFileName.lastIndexOf("."));
			String fileExt = origFileName.substring(origFileName.lastIndexOf("."), origFileName.length());
			Path targetFile = system.getPath(original.getParent() + "\\" + fileName + "_copy" + fileExt);
			targetFilePath = targetFile.toString();

			try {
				readValidateExcel.readAndValidateExcelData(tdoFilepath, cadacFilepath, cadacSheetName);
				File file = new File(targetFilePath);
				file.createNewFile();
				readValidateExcel.copyFileSheet(tdoFilepath, targetFilePath, tdoSheetName);
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
}
