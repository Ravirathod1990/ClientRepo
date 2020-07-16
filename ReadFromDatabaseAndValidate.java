import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.HashMap;
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

public class ReadFromDatabaseAndValidate {

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
			columnMap.put(cell.getStringCellValue().trim().toLowerCase(), cell.getColumnIndex());
		}

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForAgency = columnMap.get("Agency ID".toLowerCase());
				Cell agencyCell = dataRow.getCell(idxForAgency);
				String agencyId = checkCellTypeAndReturn(agencyCell);
				ConfigData configData = getConfigData(conn, agencyId);
				DacMatricesRowData dacMatricesRowData = getDacMatricesRowData(conn, agencyId);

				XSSFCell dbCell = (XSSFCell) dataRow.createCell(columnMap.get("State".toLowerCase()));
				dbCell.setCellType(CellType.STRING);
				dbCell.setCellValue(configData.getStateId());

				// Compare File count
				XSSFCell fileCount = (XSSFCell) dataRow.createCell(columnMap.get("File_count".toLowerCase()));
				fileCount.setCellType(CellType.STRING);
				fileCount.setCellValue(configData.getFileCount());
				if (!configData.getFileCount().equals(dacMatricesRowData.getFileCount())) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					fileCount.setCellStyle(style);
				}
				XSSFCell fileCountDac = (XSSFCell) dataRow.createCell(columnMap.get("File_count_DAC".toLowerCase()));
				fileCountDac.setCellType(CellType.STRING);
				fileCountDac.setCellValue(dacMatricesRowData.getFileCount());

				// Compare Process type
				XSSFCell processType = (XSSFCell) dataRow.createCell(columnMap.get("Process_type".toLowerCase()));
				processType.setCellType(CellType.STRING);
				processType.setCellValue(configData.getProcessType());
				if (!configData.getProcessType().equals(dacMatricesRowData.getProcessType())) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					processType.setCellStyle(style);
				}
				XSSFCell processTypeDac = (XSSFCell) dataRow
						.createCell(columnMap.get("Process_type_DAC".toLowerCase()));
				processTypeDac.setCellType(CellType.STRING);
				processTypeDac.setCellValue(dacMatricesRowData.getProcessType());

				// Compare File type
				XSSFCell fileType = (XSSFCell) dataRow.createCell(columnMap.get("file_type".toLowerCase()));
				fileType.setCellType(CellType.STRING);
				fileType.setCellValue(configData.getFileType());
				if (!configData.getFileType().equals(dacMatricesRowData.getFileType())) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					fileType.setCellStyle(style);
				}
				XSSFCell fileTypeDac = (XSSFCell) dataRow.createCell(columnMap.get("file_type_DAC".toLowerCase()));
				fileTypeDac.setCellType(CellType.STRING);
				fileTypeDac.setCellValue(dacMatricesRowData.getFileType());

				// Compare Record length
				XSSFCell recordLength = (XSSFCell) dataRow.createCell(columnMap.get("record_length".toLowerCase()));
				recordLength.setCellType(CellType.STRING);
				recordLength.setCellValue(configData.getRecordLength());
				if (!configData.getRecordLength().equals(dacMatricesRowData.getRecordLength())) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					recordLength.setCellStyle(style);
				}
				XSSFCell recordLengthDac = (XSSFCell) dataRow
						.createCell(columnMap.get("record_length_DAC".toLowerCase()));
				recordLengthDac.setCellType(CellType.STRING);
				recordLengthDac.setCellValue(dacMatricesRowData.getRecordLength());

				// Compare Record count
				XSSFCell recordCount = (XSSFCell) dataRow.createCell(columnMap.get("record_count".toLowerCase()));
				recordCount.setCellType(CellType.STRING);
				recordCount.setCellValue(configData.getRecordCount());
				if (!configData.getRecordCount().equals(dacMatricesRowData.getRecordCount())) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					recordCount.setCellStyle(style);
				}
				XSSFCell recordCountDac = (XSSFCell) dataRow
						.createCell(columnMap.get("record_count_DAC".toLowerCase()));
				recordCountDac.setCellType(CellType.STRING);
				recordCountDac.setCellValue(dacMatricesRowData.getRecordCount());

				// Compare Installment pos
				XSSFCell installmentPos = (XSSFCell) dataRow.createCell(columnMap.get("Installment_pos".toLowerCase()));
				installmentPos.setCellType(CellType.STRING);
				installmentPos.setCellValue(configData.getFileType());
				if (!configData.getInstallmentPos().equals(dacMatricesRowData.getLayoutInst())) {
					CellStyle style = workbook.createCellStyle();
					style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
					style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					installmentPos.setCellStyle(style);
				}
				XSSFCell installmentPosDac = (XSSFCell) dataRow
						.createCell(columnMap.get("Installment_pos_DAC".toLowerCase()));
				installmentPosDac.setCellType(CellType.STRING);
				installmentPosDac.setCellValue(dacMatricesRowData.getLayoutInst());
			} else {
				physicalNumberOfRows++;
			}
		}

		FileOutputStream fileOut = new FileOutputStream(new File(filePath));
		workbook.write(fileOut);
		workbook.close();
		fileOut.close();
	}

	public ConfigData getConfigData(Connection conn, String agencyId) {
		ConfigData configData = new ConfigData();
		Statement statement = null;
		try {
			statement = conn.createStatement();
			ResultSet resultSet = statement.executeQuery(
					"select state_id,file_count,process_type,file_type,record_length,record_count,installment_pos from config where agency_id = '"
							+ agencyId + "' and process_type='DLQ'");
			if (resultSet.next()) {
				System.out.println("data found");
				configData.setStateId(resultSet.getString(1));
				configData.setFileCount(resultSet.getString(2));
				configData.setProcessType(resultSet.getString(3));
				configData.setFileType(resultSet.getString(4));
				configData.setRecordLength(resultSet.getString(5));
				configData.setRecordCount(resultSet.getString(6));
				configData.setInstallmentPos(resultSet.getString(7));
			} else {
				System.out.println("data not found");
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

	public DacMatricesRowData getDacMatricesRowData(Connection conn, String agencyId) {
		DacMatricesRowData dacMatricesRowData = new DacMatricesRowData();
		Statement statement = null;
		try {
			statement = conn.createStatement();
			ResultSet resultSet = statement.executeQuery(
					"select file_count,process_type,file_type,record_length,record_count,layout_instructions from dac_matrices_raw where tax_authority_id = '"
							+ agencyId + "' and process_type='DLQ'");
			if (resultSet.next()) {
				System.out.println("data found");
				dacMatricesRowData.setFileCount(resultSet.getString(1));
				dacMatricesRowData.setProcessType(resultSet.getString(2));
				dacMatricesRowData.setFileType(resultSet.getString(3));
				dacMatricesRowData.setRecordLength(resultSet.getString(4));
				dacMatricesRowData.setRecordCount(resultSet.getString(5));
				dacMatricesRowData.setLayoutInst(resultSet.getString(6));
			} else {
				System.out.println("data not found");
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
		return dacMatricesRowData;
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
		String sheetName = "Sheet1";

		ReadFromDatabaseAndValidate dataClass = new ReadFromDatabaseAndValidate();
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
		private String fileCount = StringUtils.EMPTY;
		private String processType = StringUtils.EMPTY;
		private String fileType = StringUtils.EMPTY;
		private String recordLength = StringUtils.EMPTY;
		private String recordCount = StringUtils.EMPTY;
		private String installmentPos = StringUtils.EMPTY;

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

		public String getFileCount() {
			return fileCount;
		}

		public void setFileCount(String fileCount) {
			this.fileCount = fileCount;
		}

		public String getProcessType() {
			return processType;
		}

		public void setProcessType(String processType) {
			this.processType = processType;
		}

		public String getFileType() {
			return fileType;
		}

		public void setFileType(String fileType) {
			this.fileType = fileType;
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

		public String getInstallmentPos() {
			return installmentPos;
		}

		public void setInstallmentPos(String installmentPos) {
			this.installmentPos = installmentPos;
		}
	}

	public class DacMatricesRowData {
		private String taxAuthorityId = StringUtils.EMPTY;
		private String fileCount = StringUtils.EMPTY;
		private String processType = StringUtils.EMPTY;
		private String fileType = StringUtils.EMPTY;
		private String recordLength = StringUtils.EMPTY;
		private String recordCount = StringUtils.EMPTY;
		private String layoutInst = StringUtils.EMPTY;

		public String getTaxAuthorityId() {
			return taxAuthorityId;
		}

		public void setTaxAuthorityId(String taxAuthorityId) {
			this.taxAuthorityId = taxAuthorityId;
		}

		public String getFileCount() {
			return fileCount;
		}

		public void setFileCount(String fileCount) {
			this.fileCount = fileCount;
		}

		public String getProcessType() {
			return processType;
		}

		public void setProcessType(String processType) {
			this.processType = processType;
		}

		public String getFileType() {
			return fileType;
		}

		public void setFileType(String fileType) {
			this.fileType = fileType;
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

		public String getLayoutInst() {
			return layoutInst;
		}

		public void setLayoutInst(String layoutInst) {
			this.layoutInst = layoutInst;
		}
	}
}
