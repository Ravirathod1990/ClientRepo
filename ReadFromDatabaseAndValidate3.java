import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
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

public class ReadFromDatabaseAndValidate3 {

	private Map<String, Integer> columnMap = new HashMap<String, Integer>();

	public void readDbDataAndMatchAndCopyToExcel(String filePath, String sheetName) throws Exception {
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

		// New code change
		XSSFCell stateCell = (XSSFCell) row.createCell(maxColIx);
		stateCell.setCellType(CellType.STRING);
		stateCell.setCellValue("State");
		columnMap.put("State", maxColIx);
		maxColIx++;

		XSSFCell tapCell = (XSSFCell) row.createCell(maxColIx);
		tapCell.setCellType(CellType.STRING);
		tapCell.setCellValue("TAP#");
		columnMap.put("TAP#", maxColIx);
		maxColIx++;

		XSSFCell jiraRlCell = (XSSFCell) row.createCell(maxColIx);
		jiraRlCell.setCellType(CellType.STRING);
		jiraRlCell.setCellValue("JIRA_RL");
		columnMap.put("JIRA_RL", maxColIx);
		maxColIx++;

		XSSFCell jiraRcCell = (XSSFCell) row.createCell(maxColIx);
		jiraRcCell.setCellType(CellType.STRING);
		jiraRcCell.setCellValue("JIRA_RC");
		columnMap.put("JIRA_RC", maxColIx);
		maxColIx++;

		XSSFCell configDbRlCell = (XSSFCell) row.createCell(maxColIx);
		configDbRlCell.setCellType(CellType.STRING);
		configDbRlCell.setCellValue("ConfigDB_RL");
		columnMap.put("ConfigDB_RL", maxColIx);
		maxColIx++;

		XSSFCell configDbRcCell = (XSSFCell) row.createCell(maxColIx);
		configDbRcCell.setCellType(CellType.STRING);
		configDbRcCell.setCellValue("ConfigDB_RC");
		columnMap.put("ConfigDB_RC", maxColIx);
		maxColIx++;

		XSSFCell ttCell = (XSSFCell) row.createCell(maxColIx);
		ttCell.setCellType(CellType.STRING);
		ttCell.setCellValue("Tax_Tranformation");
		columnMap.put("Tax_Tranformation", maxColIx);
		maxColIx++;

		XSSFCell ttmCell = (XSSFCell) row.createCell(maxColIx);
		ttmCell.setCellType(CellType.STRING);
		ttmCell.setCellValue("TT_in_Migration Tracker");
		columnMap.put("TT_in_Migration Tracker", maxColIx);

		List<String> agencyList = new ArrayList<>();
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForAgency = columnMap.get("Agency ID");
				Cell agencyCell = dataRow.getCell(idxForAgency);
				String agencyId = checkCellTypeAndReturn(agencyCell);
				agencyList.add(agencyId);
			} else {
				physicalNumberOfRows++;
			}
		}
		Map<String, List<String>> webAgencydataList = getWebAgencyDataList(agencyList);
		// End of new code

		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForAgency = columnMap.get("Agency ID");
				Cell agencyCell = dataRow.getCell(idxForAgency);
				String agencyId = checkCellTypeAndReturn(agencyCell);
				ConfigData configData = getConfigData(conn, agencyId);

				XSSFCell dbCell = (XSSFCell) dataRow.createCell(columnMap.get("State"));
				dbCell.setCellType(CellType.STRING);
				dbCell.setCellValue(configData.getStateId());

				// -------------new code to insert data from web----------------//
				List<String> webAgencyData = webAgencydataList.get(agencyId);

				// To create TAP# cell and insert data
				XSSFCell tap = (XSSFCell) dataRow.createCell(columnMap.get("TAP#"));
				tap.setCellType(CellType.STRING);
				tap.setCellValue(webAgencyData.get(0));

				// To create JIRA_RL cell and insert data
				XSSFCell jiraRl = (XSSFCell) dataRow.createCell(columnMap.get("JIRA_RL"));
				jiraRl.setCellType(CellType.STRING);
				jiraRl.setCellValue(webAgencyData.get(1));
				if (!configData.getRecordLength().equals(webAgencyData.get(1))) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					jiraRl.setCellStyle(style);
				}

				// To create JIRA_RC cell and insert data
				XSSFCell jiraRc = (XSSFCell) dataRow.createCell(columnMap.get("JIRA_RC"));
				jiraRc.setCellType(CellType.STRING);
				jiraRc.setCellValue(webAgencyData.get(2));
				if (!configData.getRecordCount().equals(webAgencyData.get(2))) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					jiraRc.setCellStyle(style);
				}

				// Compare Record length
				XSSFCell recordLength = (XSSFCell) dataRow.createCell(columnMap.get("ConfigDB_RL"));
				recordLength.setCellType(CellType.STRING);
				recordLength.setCellValue(configData.getRecordLength());
				if (!configData.getRecordLength().equals(webAgencyData.get(1))) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					recordLength.setCellStyle(style);
				}

				// Compare Record count
				XSSFCell recordCount = (XSSFCell) dataRow.createCell(columnMap.get("ConfigDB_RC"));
				recordCount.setCellType(CellType.STRING);
				recordCount.setCellValue(configData.getRecordCount());
				if (!configData.getRecordCount().equals(webAgencyData.get(2))) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					recordCount.setCellStyle(style);
				}
				XSSFCell tt = (XSSFCell) dataRow.createCell(columnMap.get("Tax_Tranformation"));
				tt.setCellType(CellType.STRING);
				tt.setCellValue(configData.getTt());

				XSSFCell ttm = (XSSFCell) dataRow.createCell(columnMap.get("TT_in_Migration Tracker"));
				ttm.setCellType(CellType.STRING);
				ttm.setCellValue(configData.getTtm());
			} else {
				physicalNumberOfRows++;
			}
		}

		FileOutputStream fileOut = new FileOutputStream(new File(filePath));
		workbook.write(fileOut);
		workbook.close();
		fileOut.close();
	}

	// New method to fetch data from web API
	public Map<String, List<String>> getWebAgencyDataList(List<String> agencyList) {
		Map<String, List<String>> webAgencydataList = new HashMap<String, List<String>>();
		for (String agencyId : agencyList) {

			// API to fetch data and put into map
			List<String> agencyDataList = new ArrayList<String>();
			agencyDataList.add("4"); // TAP#
			agencyDataList.add("50"); // JIRA_RL
			agencyDataList.add("15"); // JIRA_RC
			webAgencydataList.put(agencyId, agencyDataList);
		}
		return webAgencydataList;
	}

	public ConfigData getConfigData(Connection conn, String agencyId) {
		ConfigData configData = new ConfigData();
		Statement statement = null;
		try {
			statement = conn.createStatement();
			ResultSet rs1 = statement
					.executeQuery("select state_id,record_length,record_count from config where agency_id = '"
							+ agencyId + "' and process_type='DLQ'");
			if (rs1.next()) {
				System.out.println("data found from config table");
				configData.setStateId(rs1.getString(1));
				configData.setRecordLength(rs1.getString(2));
				configData.setRecordCount(rs1.getString(3));
			} else {
				System.out.println("data not found from config table");
			}

			ResultSet rs2 = statement
					.executeQuery("select tax_transformation from dac_matrices_raw where tax_authority_id = '"
							+ agencyId + "' and process_type='DLQ'");
			if (rs2.next()) {
				System.out.println("data found from dac_matrices_raw table");
				configData.setTt(rs2.getString(1));
			} else {
				System.out.println("data not found from dac_matrices_raw table");
			}

			if (!configData.getTt().equals(StringUtils.EMPTY)) {
				ResultSet rs3 = statement
						.executeQuery("select tax_transformation from viper_raw where tax_transformation = '"
								+ configData.getTt() + "'");
				if (rs3.next()) {
					System.out.println("data found from viper_raw table");
					configData.setTtm("Exist");
				} else {
					configData.setTtm("Not Exist");
					System.out.println("data not found from viper_raw table");
				}
			} else {
				configData.setTtm("Not Exist");
			}
		} catch (Exception ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		} finally {
			try {
				if (statement != null) {
					statement.close();
				}
			} catch (SQLException e) {
				e.printStackTrace();
			}
		}
		return configData;
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

		String filePath = "E:\\client\\files\\new\\Compare.xlsx";
		String sheetName = "Sheet2";

		ReadFromDatabaseAndValidate3 dataClass = new ReadFromDatabaseAndValidate3();
		dataClass.readDbDataAndMatchAndCopyToExcel(filePath, sheetName);

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

	public class ConfigData {
		private String agencyId = StringUtils.EMPTY;
		private String stateId = StringUtils.EMPTY;
		private String recordLength = StringUtils.EMPTY;
		private String recordCount = StringUtils.EMPTY;
		private String tt = StringUtils.EMPTY;
		private String ttm = StringUtils.EMPTY;

		public String getAgencyId() {
			return agencyId;
		}

		public void setAgencyId(String agencyId) {
			this.agencyId = agencyId;
		}

		public String getStateId() {
			return stateId;
		}

		public void setStateId(String stateId) {
			this.stateId = stateId;
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

		public String getTt() {
			return tt;
		}

		public void setTt(String tt) {
			this.tt = tt;
		}

		public String getTtm() {
			return ttm;
		}

		public void setTtm(String ttm) {
			this.ttm = ttm;
		}
	}
}
