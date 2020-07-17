import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadAndValidateExcelWithDatabase6 {

	private List<ProfileRow> dataList = new ArrayList<ProfileRow>();
	private Map<String, Integer> columnMap = new HashMap<String, Integer>();

	public void readExcelAndInsertDataToDB(String filePath, String sheetName) throws Exception {
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
				int idxForStatus = columnMap.get("Status");

				Cell agencyCell = dataRow.getCell(idxForAgency);
				Cell trspCell = dataRow.getCell(idxForTrsp);
				Cell taxTransCell = dataRow.getCell(idxForTaxTrans);
				Cell qmeFilterCell = dataRow.getCell(idxForQmeFilter);
				Cell specInstCell = dataRow.getCell(idxForSpecInst);
				Cell paymentCell = dataRow.getCell(idxForPayment);
				Cell skipUpdateCell = dataRow.getCell(idxForSkipUpdate);
				Cell dacNoteCell = dataRow.getCell(idxForDacNote);
				Cell statusCell = dataRow.getCell(idxForStatus);

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
				if (statusCell != null) {
					profileRow.setStatus(checkCellTypeAndReturn(statusCell));
				}

				if (profileRow.getAgency() != null) {
					dataList.add(profileRow);
				}
			} else {
				physicalNumberOfRows++;
			}
		}

		for (ProfileRow profileRow : dataList) {
			if (profileRow.getStatus().equals("Data not available in database")) {
				// Insert query for database
				PreparedStatement stmt = conn.prepareStatement(
						"INSERT INTO profile (agency_id,trsp_recl,tax_transformation,qme_execution,other_special_instr,payments,"
								+ "skip_updates,dac_notes) values(?,?,?,?,?,?,?,?)");
				stmt.setString(1, profileRow.getAgency());
				stmt.setString(2, profileRow.getTrsp());
				stmt.setString(3, profileRow.getTaxTran());
				stmt.setString(4, profileRow.getQmeFilter());
				stmt.setString(5, profileRow.getSpecialInst());
				stmt.setString(6, profileRow.getPayment());
				stmt.setString(7, profileRow.getSkipUpdate());
				stmt.setString(8, profileRow.getDacNote());
				stmt.execute();
			} else if (!profileRow.getStatus().equals("Passed")) {
				// Update query for database
				PreparedStatement stmt = conn
						.prepareStatement("UPDATE profile set agency_id = ?, trsp_recl = ?, tax_transformation = ?, "
								+ "qme_execution = ?, other_special_instr = ?, payments = ?, skip_updates = ?, "
								+ "dac_notes = ? where agency_id = ? and tax_transformation = ?");
				stmt.setString(1, profileRow.getAgency());
				stmt.setString(2, profileRow.getTrsp());
				stmt.setString(3, profileRow.getTaxTran());
				stmt.setString(4, profileRow.getQmeFilter());
				stmt.setString(5, profileRow.getSpecialInst());
				stmt.setString(6, profileRow.getPayment());
				stmt.setString(7, profileRow.getSkipUpdate());
				stmt.setString(8, profileRow.getDacNote());
				stmt.setString(9, profileRow.getAgency());
				stmt.setString(10, profileRow.getTaxTran());
				stmt.executeUpdate();
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

	public static void main(String[] args) throws Exception {

		String filePath = "E:\\client\\files\\newexcel\\Profile_copy.xlsx";
		String sheetName = "Sheet1";

		ReadAndValidateExcelWithDatabase6 dataClass = new ReadAndValidateExcelWithDatabase6();
		dataClass.readExcelAndInsertDataToDB(filePath, sheetName);

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

		private String status = StringUtils.EMPTY;

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

		public String getStatus() {
			return status;
		}

		public void setStatus(String status) {
			this.status = status;
		}
	}
}
