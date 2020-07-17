import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndValidateExcelWithDatabase5 {

	private Map<String, ProfileRow> dataList = new LinkedHashMap<String, ProfileRow>();
	private Map<String, Integer> columnMap = new HashMap<String, Integer>();

	public void readExcel(String filePath, String sheetName) throws Exception {
		Connection conn = createDBConnection();

		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			columnMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			ProfileRow profileRow = new ProfileRow();
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForAgency = columnMap.get("agency#");
				int idxForTrsp = columnMap.get("TRSP");
				int idxForTaxTrans = columnMap.get("TAXTRANSFORMATION");
				int idxForQmeFilter = columnMap.get("QME Filter");
				int idxForSpecInst = columnMap.get("other special Instructions");
				int idxForPayment = columnMap.get("Payments");
				int idxForSkipUpdate = columnMap.get("skip updates");
				int idxForDacNote = columnMap.get("dac_notes");

				Cell agencyCell = dataRow.getCell(idxForAgency);
				Cell trspCell = dataRow.getCell(idxForTrsp);
				Cell taxTransCell = dataRow.getCell(idxForTaxTrans);
				Cell qmeFilterCell = dataRow.getCell(idxForQmeFilter);
				Cell specInstCell = dataRow.getCell(idxForSpecInst);
				Cell paymentCell = dataRow.getCell(idxForPayment);
				Cell skipUpdateCell = dataRow.getCell(idxForSkipUpdate);
				Cell dacNoteCell = dataRow.getCell(idxForDacNote);

				if (agencyCell != null) {
					profileRow.setAgency(checkCellTypeAndReturn(agencyCell));
				}
				if (trspCell != null) {
					profileRow.setTrsp(checkCellTypeAndReturn(trspCell));
				}
				if (taxTransCell != null) {
					profileRow.setTaxTran(checkCellTypeAndReturn(taxTransCell));
				}
				if (qmeFilterCell != null) {
					profileRow.setQmeFilter(checkCellTypeAndReturn(qmeFilterCell));
				}
				if (specInstCell != null) {
					profileRow.setSpecialInst(checkCellTypeAndReturn(specInstCell));
				}
				if (paymentCell != null) {
					profileRow.setPayment(checkCellTypeAndReturn(paymentCell));
				}
				if (skipUpdateCell != null) {
					profileRow.setSkipUpdate(checkCellTypeAndReturn(skipUpdateCell));
				}
				if (dacNoteCell != null) {
					profileRow.setDacNote(checkCellTypeAndReturn(dacNoteCell));
				}

				if (profileRow.getAgency() != null) {
					getDbRecordCount(conn, profileRow);
					dataList.put(profileRow.getAgency() + "_" + profileRow.getTaxTran(), profileRow);
				}
			} else {
				physicalNumberOfRows++;
			}
		}
	}

	public List<String> getDbRecordCount(Connection conn, ProfileRow profileRow) {
		List<String> dbRecordCountList = new ArrayList<>();

		try {
			Statement statement = conn.createStatement();
			ResultSet resultSet = statement.executeQuery(
					"select agency_id,tax_transformation,trsp_recl,qme_execution,other_special_instr,payments,skip_updates,dac_notes from profile where agency_id = '"
							+ profileRow.getAgency() + "' and tax_transformation = '" + profileRow.getTaxTran() + "' ");
			if (resultSet.next()) {
				System.out.println("data found");
				boolean passed = true;
				StringBuffer sbStatus = new StringBuffer();
				profileRow.setDbAgency(resultSet.getString(1));
				profileRow.setDbTaxTran(resultSet.getString(2));
				profileRow.setDbTrsp(resultSet.getString(3));
				if (!profileRow.getTrsp().equals(profileRow.getDbTrsp())) {
					profileRow.getDataNotMatchList().add("TRSP");
					sbStatus.append("TRSP not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				profileRow.setDbQmeFilter(resultSet.getString(4));
				if (!profileRow.getQmeFilter().equals(profileRow.getDbQmeFilter())) {
					profileRow.getDataNotMatchList().add("QME Filter");
					sbStatus.append("QME Filter not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				profileRow.setDbSpecialInst(resultSet.getString(5));
				if (!profileRow.getSpecialInst().equals(profileRow.getDbSpecialInst())) {
					profileRow.getDataNotMatchList().add("other special Instructions");
					sbStatus.append("other special Instructions not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				profileRow.setDbPayment(resultSet.getString(6));
				if (!profileRow.getPayment().equals(profileRow.getDbPayment())) {
					profileRow.getDataNotMatchList().add("Payments");
					sbStatus.append("Payments not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				profileRow.setDbSkipUpdate(resultSet.getString(7));
				if (!profileRow.getSkipUpdate().equals(profileRow.getDbSkipUpdate())) {
					profileRow.getDataNotMatchList().add("skip updates");
					sbStatus.append("skip updates not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				profileRow.setDbDacNote(resultSet.getString(8));
				if (!profileRow.getDacNote().equals(profileRow.getDbDacNote())) {
					profileRow.getDataNotMatchList().add("dac_notes");
					sbStatus.append("dac notes not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				if (passed) {
					profileRow.setStatus("Passed");
				}
				if (!passed) {
					profileRow.setStatus(sbStatus.toString());
				}
			} else {
				System.out.println("data not found");
				profileRow.setStatus("Data not available in database");
			}
		} catch (Exception ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		}

		return dbRecordCountList;
	}

	public void copyFileSheet(String filePath, String targetFilePath, String sheetName, String dbSheetName)
			throws IOException {
		BufferedInputStream bis = new BufferedInputStream(new FileInputStream(filePath));
		XSSFWorkbook workbook = new XSSFWorkbook(bis);
		XSSFWorkbook myWorkBook = new XSSFWorkbook();
		XSSFSheet sheet = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		XSSFSheet mySheet = null;
		XSSFSheet dbSheet = null;
		XSSFRow myRow = null;
		XSSFRow dbRow = null;
		XSSFCell myCell = null;
		XSSFCell dbCell = null;
		int fCell = 0;
		int lCell = 0;
		int headerlCell = 0;
		int fRow = 0;
		int lRow = 0;

		sheet = workbook.getSheet(sheetName);
		if (sheet != null) {
			mySheet = myWorkBook.createSheet(sheet.getSheetName());
			dbSheet = myWorkBook.createSheet(dbSheetName);
			fRow = sheet.getFirstRowNum();
			lRow = sheet.getLastRowNum();
			for (int iRow = fRow; iRow <= lRow; iRow++) {
				row = sheet.getRow(iRow);
				myRow = mySheet.createRow(iRow);
				dbRow = dbSheet.createRow(iRow);
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
						}
					}
					if (iRow == 0) {
						myCell = myRow.createCell(lCell);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue("Status");

						List<String> list = new ArrayList<>();
						list.add("DB agency");
						list.add("DB TRSP");
						list.add("DB TAXTRANSFORMATION");
						list.add("DB QME Filter");
						list.add("DB other special Instructions");
						list.add("DB Payments");
						list.add("DB skip updates");
						list.add("DB dac_notes");

						int counter = 0;
						for (int iCell = fCell; iCell < lCell; iCell++) {
							dbCell = dbRow.createCell(iCell);
							dbCell.setCellType(CellType.STRING);
							dbCell.setCellValue(list.get(counter));
							counter++;
						}
					} else {
						Cell agencyCell = row.getCell(columnMap.get("agency#"));
						String agencyValue = checkCellTypeAndReturn(agencyCell);

						Cell taxTransCell = row.getCell(columnMap.get("TAXTRANSFORMATION"));
						String taxTransValue = checkCellTypeAndReturn(taxTransCell);

						ProfileRow profileRow = dataList.get(agencyValue + "_" + taxTransValue);
						int counter = 0;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbAgency());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbTrsp());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbTaxTran());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbQmeFilter());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbSpecialInst());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbPayment());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbSkipUpdate());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(profileRow.getDbDacNote());

						myCell = myRow.createCell(lCell);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(profileRow.getStatus());

						for (String columnName : profileRow.getDataNotMatchList()) {
							int idxForColumnName = columnMap.get(columnName);
							Cell myRowCell = myRow.getCell(idxForColumnName);
							CellStyle style = myWorkBook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							myRowCell.setCellStyle(style);

							Cell dbRowCell = dbRow.getCell(idxForColumnName);
							dbRowCell.setCellStyle(style);
						}
					}
				}
			}
		}
		FileOutputStream fileOut = new FileOutputStream(new File(targetFilePath));
		myWorkBook.write(fileOut);
		bis.close();
		workbook.close();
		myWorkBook.close();
		fileOut.close();
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

	public static void main(String[] args) throws Exception {

		String filePath = "E:\\client\\files\\newexcel\\Profile.xlsx";
		String sheetName = "Sheet1";
		String dbSheetName = "DB";
		String targetFilePath;

		FileSystem system = FileSystems.getDefault();
		Path original = system.getPath(filePath);
		String origFileName = original.getFileName().toString();
		String fileName = origFileName.substring(0, origFileName.lastIndexOf("."));
		String fileExt = origFileName.substring(origFileName.lastIndexOf("."), origFileName.length());
		Path targetFile = system.getPath(original.getParent() + "\\" + fileName + "_copy" + fileExt);
		targetFilePath = targetFile.toString();

		ReadAndValidateExcelWithDatabase5 dataClass = new ReadAndValidateExcelWithDatabase5();
		dataClass.readExcel(filePath, sheetName);

		File file = new File(targetFilePath);
		file.createNewFile();
		dataClass.copyFileSheet(filePath, targetFilePath, sheetName, dbSheetName);

		System.out.println("Process completed successfully");
	}

	public Connection createDBConnection() {
		Connection conn = null;
		try {
			String url = "jdbc:mysql://localhost:3306/test?user=root&password=";
			conn = DriverManager.getConnection(url);
		} catch (SQLException ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		}
		return conn;
	}

	public class ProfileRow {
		private String agency = StringUtils.EMPTY;
		private String trsp = StringUtils.EMPTY;
		private String taxTran = StringUtils.EMPTY;
		private String qmeFilter = StringUtils.EMPTY;
		private String specialInst = StringUtils.EMPTY;
		private String payment = StringUtils.EMPTY;
		private String skipUpdate = StringUtils.EMPTY;
		private String dacNote = StringUtils.EMPTY;

		private String dbAgency = StringUtils.EMPTY;
		private String dbTrsp = StringUtils.EMPTY;
		private String dbTaxTran = StringUtils.EMPTY;
		private String dbQmeFilter = StringUtils.EMPTY;
		private String dbSpecialInst = StringUtils.EMPTY;
		private String dbPayment = StringUtils.EMPTY;
		private String dbSkipUpdate = StringUtils.EMPTY;
		private String dbDacNote = StringUtils.EMPTY;

		private String status = StringUtils.EMPTY;
		private List<String> dataNotMatchList = new ArrayList<String>();

		public String getAgency() {
			return agency;
		}

		public void setAgency(String agency) {
			this.agency = agency;
		}

		public String getTrsp() {
			return trsp;
		}

		public void setTrsp(String trsp) {
			this.trsp = trsp;
		}

		public String getTaxTran() {
			return taxTran;
		}

		public void setTaxTran(String taxTran) {
			this.taxTran = taxTran;
		}

		public String getQmeFilter() {
			return qmeFilter;
		}

		public void setQmeFilter(String qmeFilter) {
			this.qmeFilter = qmeFilter;
		}

		public String getSpecialInst() {
			return specialInst;
		}

		public void setSpecialInst(String specialInst) {
			this.specialInst = specialInst;
		}

		public String getPayment() {
			return payment;
		}

		public void setPayment(String payment) {
			this.payment = payment;
		}

		public String getSkipUpdate() {
			return skipUpdate;
		}

		public void setSkipUpdate(String skipUpdate) {
			this.skipUpdate = skipUpdate;
		}

		public String getDacNote() {
			return dacNote;
		}

		public void setDacNote(String dacNote) {
			this.dacNote = dacNote;
		}

		public String getDbAgency() {
			return dbAgency;
		}

		public void setDbAgency(String dbAgency) {
			this.dbAgency = dbAgency;
		}

		public String getDbTrsp() {
			return dbTrsp;
		}

		public void setDbTrsp(String dbTrsp) {
			this.dbTrsp = dbTrsp;
		}

		public String getDbTaxTran() {
			return dbTaxTran;
		}

		public void setDbTaxTran(String dbTaxTran) {
			this.dbTaxTran = dbTaxTran;
		}

		public String getDbQmeFilter() {
			return dbQmeFilter;
		}

		public void setDbQmeFilter(String dbQmeFilter) {
			this.dbQmeFilter = dbQmeFilter;
		}

		public String getDbSpecialInst() {
			return dbSpecialInst;
		}

		public void setDbSpecialInst(String dbSpecialInst) {
			this.dbSpecialInst = dbSpecialInst;
		}

		public String getDbPayment() {
			return dbPayment;
		}

		public void setDbPayment(String dbPayment) {
			this.dbPayment = dbPayment;
		}

		public String getDbSkipUpdate() {
			return dbSkipUpdate;
		}

		public void setDbSkipUpdate(String dbSkipUpdate) {
			this.dbSkipUpdate = dbSkipUpdate;
		}

		public String getDbDacNote() {
			return dbDacNote;
		}

		public void setDbDacNote(String dbDacNote) {
			this.dbDacNote = dbDacNote;
		}

		public String getStatus() {
			return status;
		}

		public void setStatus(String status) {
			this.status = status;
		}

		public List<String> getDataNotMatchList() {
			return dataNotMatchList;
		}

		public void setDataNotMatchList(List<String> dataNotMatchList) {
			this.dataNotMatchList = dataNotMatchList;
		}
	}
}
