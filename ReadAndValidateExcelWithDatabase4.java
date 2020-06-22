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

public class ReadAndValidateExcelWithDatabase4 {

	private List<ViperRow> dataList = new ArrayList<ViperRow>(); // get data from excel and store into list
	private Map<String, Integer> columnMap = new HashMap<String, Integer>(); // To fetch header name and index from sheet and store into map

	public void readExcelAndInsertDataToDB(String filePath, String sheetName) throws Exception {
		Connection conn = createDBConnection();

		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));
		Sheet sheet = workbook.getSheet(sheetName); // To get sheet from the excel file
		Row row = sheet.getRow(0); // To get header row from the sheet

		int minColIx = row.getFirstCellNum(); // To get first cell number from the row (ex. 1 for all cases it's 1 minimum)
		int maxColIx = row.getLastCellNum(); // To get last cell number from the row (ex. 25 last cell number)

		for (int colIx = minColIx; colIx < maxColIx; colIx++) { // iterate through all cell in header row and store cell value and it's index into map
			Cell cell = row.getCell(colIx);
			columnMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) { // iterate through all rows from sheet and store into list
			ViperRow viperRow = new ViperRow();
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForTt = columnMap.get("TT");
				int idxForTitle = columnMap.get("Title");
				int idxForVendor = columnMap.get("Vendor");
				int idxForInProduction = columnMap.get("In Production");
				int idxForEnabledDisabled = columnMap.get("Enabled/Disabled");
				int idxForFileBillYearTyep = columnMap.get("File Bill Year Type");
				int idxForFilePaymentStatusType = columnMap.get("File Payment Status Type");
				int idxForFileBillType = columnMap.get("File Bill Type");
				int idxForTimingConstraints = columnMap.get("Timing Constrains");
				int idxForCommentForProcure = columnMap.get("Comment for Procure");
				int idxForCommentProces = columnMap.get("Comment Processor");
				int idxForHardCodes = columnMap.get("Hard codes");
				int idxForRscs = columnMap.get("RSC(s)");
				int idxForState = columnMap.get("State(s)");
				int idxForTaxTransformCount = columnMap.get("Tax Transform Count");
				int idxForTtaCount = columnMap.get("TTA Count");
				int idxForComplexityProjection = columnMap.get("Complexity Projection");
				int idxForComment = columnMap.get("Comment");
				int idxForDateSubmitJira = columnMap.get("Date Submitted JIRA");
				int idxForMigrationComplDate = columnMap.get("Migration Completion Date");
				int idxForFileDeficiences = columnMap.get("File Deficiences");
				int idxForFutureTtgEnhcNote = columnMap.get("Future TTG Enhancement Needs");
				int idxForRscPair = columnMap.get("RSC pair");
				int idxForReasonNtStage = columnMap.get("Reason not Staged");
				int idxForStatus = columnMap.get("Status");

				Cell ttCell = dataRow.getCell(idxForTt);
				Cell titleCell = dataRow.getCell(idxForTitle);
				Cell vendorCell = dataRow.getCell(idxForVendor);
				Cell inProductionCell = dataRow.getCell(idxForInProduction);
				Cell enabledDisabledCell = dataRow.getCell(idxForEnabledDisabled);
				Cell fileBillYearTyepCell = dataRow.getCell(idxForFileBillYearTyep);
				Cell filePaymentStatusTypeCell = dataRow.getCell(idxForFilePaymentStatusType);
				Cell fileBillTypeCell = dataRow.getCell(idxForFileBillType);
				Cell timingConstraintsCell = dataRow.getCell(idxForTimingConstraints);
				Cell commentForProcureCell = dataRow.getCell(idxForCommentForProcure);
				Cell commentProcesCell = dataRow.getCell(idxForCommentProces);
				Cell stateCell = dataRow.getCell(idxForState);
				Cell taxTransformCountCell = dataRow.getCell(idxForTaxTransformCount);
				Cell ttaCountCell = dataRow.getCell(idxForTtaCount);
				Cell complexityProjectionCell = dataRow.getCell(idxForComplexityProjection);
				Cell commentCell = dataRow.getCell(idxForComment);
				Cell dateSubmitJiraCell = dataRow.getCell(idxForDateSubmitJira);
				Cell migrationComplDateCell = dataRow.getCell(idxForMigrationComplDate);
				Cell fileDeficiencesCell = dataRow.getCell(idxForFileDeficiences);
				Cell futureTtgEnhcNoteCell = dataRow.getCell(idxForFutureTtgEnhcNote);
				Cell rscPairCell = dataRow.getCell(idxForRscPair);
				Cell reasonNtStageCell = dataRow.getCell(idxForReasonNtStage);
				Cell hardCodesCell = dataRow.getCell(idxForHardCodes);
				Cell rscsCell = dataRow.getCell(idxForRscs);
				Cell statusCell = dataRow.getCell(idxForStatus);

				if (ttCell != null) {
					viperRow.setTt(checkCellTypeAndReturn(ttCell));
				}
				if (titleCell != null) {
					viperRow.setTitle(checkCellTypeAndReturn(titleCell));
				}
				if (vendorCell != null) {
					viperRow.setVendor(checkCellTypeAndReturn(vendorCell));
				}
				if (inProductionCell != null) {
					viperRow.setInProduction(checkCellTypeAndReturn(inProductionCell));
				}
				if (enabledDisabledCell != null) {
					viperRow.setEnabledDisabled(checkCellTypeAndReturn(enabledDisabledCell));
				}
				if (fileBillYearTyepCell != null) {
					viperRow.setFileBillYearTyep(checkCellTypeAndReturn(fileBillYearTyepCell));
				}
				if (filePaymentStatusTypeCell != null) {
					viperRow.setFilePaymentStatusType(checkCellTypeAndReturn(filePaymentStatusTypeCell));
				}
				if (fileBillTypeCell != null) {
					viperRow.setFileBillType(checkCellTypeAndReturn(fileBillTypeCell));
				}
				if (timingConstraintsCell != null) {
					viperRow.setTimingConstraints(checkCellTypeAndReturn(timingConstraintsCell));
				}
				if (commentForProcureCell != null) {
					viperRow.setCommentForProcure(checkCellTypeAndReturn(commentForProcureCell));
				}
				if (commentProcesCell != null) {
					viperRow.setCommentProces(checkCellTypeAndReturn(commentProcesCell));
				}
				if (hardCodesCell != null) {
					viperRow.setHardCodes(checkCellTypeAndReturn(hardCodesCell));
				}
				if (rscsCell != null) {
					viperRow.setRscs(checkCellTypeAndReturn(rscsCell));
				}
				if (stateCell != null) {
					viperRow.setState(checkCellTypeAndReturn(stateCell));
				}
				if (taxTransformCountCell != null) {
					viperRow.setTaxTransformCount(checkCellTypeAndReturn(taxTransformCountCell));
				}
				if (ttaCountCell != null) {
					viperRow.setTtaCount(checkCellTypeAndReturn(ttaCountCell));
				}
				if (complexityProjectionCell != null) {
					viperRow.setComplexityProjection(checkCellTypeAndReturn(complexityProjectionCell));
				}
				if (commentCell != null) {
					viperRow.setComment(checkCellTypeAndReturn(commentCell));
				}
				if (dateSubmitJiraCell != null) {
					viperRow.setDateSubmitJira(checkCellTypeAndReturn(dateSubmitJiraCell));
				}
				if (migrationComplDateCell != null) {
					viperRow.setMigrationComplDate(checkCellTypeAndReturn(migrationComplDateCell));
				}
				if (fileDeficiencesCell != null) {
					viperRow.setFileDeficiences(checkCellTypeAndReturn(fileDeficiencesCell));
				}
				if (futureTtgEnhcNoteCell != null) {
					viperRow.setFutureTtgEnhcNote(checkCellTypeAndReturn(futureTtgEnhcNoteCell));
				}
				if (rscPairCell != null) {
					viperRow.setRscPair(checkCellTypeAndReturn(rscPairCell));
				}
				if (reasonNtStageCell != null) {
					viperRow.setReasonNtStage(checkCellTypeAndReturn(reasonNtStageCell));
				}
				if (statusCell != null) {
					viperRow.setStatus(checkCellTypeAndReturn(statusCell));
				}

				if (viperRow.getTitle() != null) {
					dataList.add(viperRow);
				}
			} else {
				physicalNumberOfRows++;
			}
		}

		for (ViperRow viperRow : dataList) { // Iterate through List of rows
			if (viperRow.getStatus().equals("Data not available in database")) { // To check whether data is exist in database (if not exist then insert otherwise update)
				// Insert query for database
				PreparedStatement stmt = conn.prepareStatement(
						"INSERT INTO vipertable (title,tax_transformation,vendor,in_production,enabled_disabled,file_bill_year_type,"
								+ "file_payment_status_type,file_bill_type,timing_constraints,comment_for_procure,comment_for_processor,"
								+ "hard_codes,rsc,states,tax_transform_count,tta_count,complexity_projection,comment,date_submitted_to_jira,"
								+ "migration_completion_date,file_deficience,future_ttg_enhancement_needs,rsc_pair,reason_not_staged) "
								+ "values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
				stmt.setString(1, viperRow.getTitle());
				stmt.setString(2, viperRow.getTt());
				stmt.setString(3, viperRow.getVendor());
				stmt.setString(4, viperRow.getInProduction());
				stmt.setString(5, viperRow.getEnabledDisabled());
				stmt.setString(6, viperRow.getFileBillYearTyep());
				stmt.setString(7, viperRow.getFilePaymentStatusType());
				stmt.setString(8, viperRow.getFileBillType());
				stmt.setString(9, viperRow.getTimingConstraints());
				stmt.setString(10, viperRow.getCommentForProcure());
				stmt.setString(11, viperRow.getCommentProces());
				stmt.setString(12, viperRow.getHardCodes());
				stmt.setString(13, viperRow.getRscs());
				stmt.setString(14, viperRow.getState());
				stmt.setString(15, viperRow.getTaxTransformCount());
				stmt.setString(16, viperRow.getTtaCount());
				stmt.setString(17, viperRow.getComplexityProjection());
				stmt.setString(18, viperRow.getComment());
				stmt.setString(19, viperRow.getDateSubmitJira());
				stmt.setString(20, viperRow.getMigrationComplDate());
				stmt.setString(21, viperRow.getFileDeficiences());
				stmt.setString(22, viperRow.getFutureTtgEnhcNote());
				stmt.setString(23, viperRow.getRscPair());
				stmt.setString(24, viperRow.getReasonNtStage());
				stmt.execute();
			} else if (!viperRow.getStatus().equals("Passed")) { // // To check data into db is exist but not updated
				// Update query for database
				PreparedStatement stmt = conn.prepareStatement(
						"UPDATE vipertable set tax_transformation = ?, vendor = ?, in_production = ?, "
								+ "enabled_disabled = ?, file_bill_year_type = ?, file_payment_status_type = ?, file_bill_type = ?, "
								+ "timing_constraints = ?, comment_for_procure = ?, comment_for_processor = ?, hard_codes = ?, "
								+ "rsc = ?, states = ?, tax_transform_count = ?, tta_count = ?, "
								+ "complexity_projection = ?, comment = ?, date_submitted_to_jira = ?, migration_completion_date = ?, "
								+ "file_deficience = ?, future_ttg_enhancement_needs = ?, "
								+ "rsc_pair = ?, reason_not_staged = ? where title = ?");
				stmt.setString(1, viperRow.getTt());
				stmt.setString(2, viperRow.getVendor());
				stmt.setString(3, viperRow.getInProduction());
				stmt.setString(4, viperRow.getEnabledDisabled());
				stmt.setString(5, viperRow.getFileBillYearTyep());
				stmt.setString(6, viperRow.getFilePaymentStatusType());
				stmt.setString(7, viperRow.getFileBillType());
				stmt.setString(8, viperRow.getTimingConstraints());
				stmt.setString(9, viperRow.getCommentForProcure());
				stmt.setString(10, viperRow.getCommentProces());
				stmt.setString(11, viperRow.getHardCodes());
				stmt.setString(12, viperRow.getRscs());
				stmt.setString(13, viperRow.getState());
				stmt.setString(14, viperRow.getTaxTransformCount());
				stmt.setString(15, viperRow.getTtaCount());
				stmt.setString(16, viperRow.getComplexityProjection());
				stmt.setString(17, viperRow.getComment());
				stmt.setString(18, viperRow.getDateSubmitJira());
				stmt.setString(19, viperRow.getMigrationComplDate());
				stmt.setString(20, viperRow.getFileDeficiences());
				stmt.setString(21, viperRow.getFutureTtgEnhcNote());
				stmt.setString(22, viperRow.getRscPair());
				stmt.setString(23, viperRow.getReasonNtStage());
				stmt.setString(24, viperRow.getTitle());
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

		String filePath = "E:\\client\\files\\newexcel\\Viper_copy.xlsx";
		String sheetName = "Sheet1";

		ReadAndValidateExcelWithDatabase4 dataClass = new ReadAndValidateExcelWithDatabase4();
		dataClass.readExcelAndInsertDataToDB(filePath, sheetName);

		System.out.println("Process completed successfully");
	}

	public Connection createDBConnection() {
		Connection conn = null;
		try {
			String url = "jdbc:mysql://localhost:3306/jpademo?user=root&password=";
			conn = DriverManager.getConnection(url);
		} catch (SQLException ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		}
		return conn;
	}

	public class ViperRow {
		private String tt = StringUtils.EMPTY;
		private String title = StringUtils.EMPTY;
		private String vendor = StringUtils.EMPTY;
		private String inProduction = StringUtils.EMPTY;
		private String enabledDisabled = StringUtils.EMPTY;
		private String fileBillYearTyep = StringUtils.EMPTY;
		private String filePaymentStatusType = StringUtils.EMPTY;
		private String fileBillType = StringUtils.EMPTY;
		private String timingConstraints = StringUtils.EMPTY;
		private String commentForProcure = StringUtils.EMPTY;
		private String commentProces = StringUtils.EMPTY;
		private String hardCodes = StringUtils.EMPTY;
		private String rscs = StringUtils.EMPTY;
		private String state = StringUtils.EMPTY;
		private String taxTransformCount = StringUtils.EMPTY;
		private String ttaCount = StringUtils.EMPTY;
		private String complexityProjection = StringUtils.EMPTY;
		private String comment = StringUtils.EMPTY;
		private String dateSubmitJira = StringUtils.EMPTY;
		private String migrationComplDate = StringUtils.EMPTY;
		private String fileDeficiences = StringUtils.EMPTY;
		private String futureTtgEnhcNote = StringUtils.EMPTY;
		private String rscPair = StringUtils.EMPTY;
		private String reasonNtStage = StringUtils.EMPTY;

		private String status = StringUtils.EMPTY;

		public String getTt() {
			return tt;
		}

		public void setTt(String tt) {
			this.tt = tt;
		}

		public String getTitle() {
			return title;
		}

		public void setTitle(String title) {
			this.title = title;
		}

		public String getVendor() {
			return vendor;
		}

		public void setVendor(String vendor) {
			this.vendor = vendor;
		}

		public String getInProduction() {
			return inProduction;
		}

		public void setInProduction(String inProduction) {
			this.inProduction = inProduction;
		}

		public String getEnabledDisabled() {
			return enabledDisabled;
		}

		public void setEnabledDisabled(String enabledDisabled) {
			this.enabledDisabled = enabledDisabled;
		}

		public String getFileBillYearTyep() {
			return fileBillYearTyep;
		}

		public void setFileBillYearTyep(String fileBillYearTyep) {
			this.fileBillYearTyep = fileBillYearTyep;
		}

		public String getFilePaymentStatusType() {
			return filePaymentStatusType;
		}

		public void setFilePaymentStatusType(String filePaymentStatusType) {
			this.filePaymentStatusType = filePaymentStatusType;
		}

		public String getFileBillType() {
			return fileBillType;
		}

		public void setFileBillType(String fileBillType) {
			this.fileBillType = fileBillType;
		}

		public String getTimingConstraints() {
			return timingConstraints;
		}

		public void setTimingConstraints(String timingConstraints) {
			this.timingConstraints = timingConstraints;
		}

		public String getCommentForProcure() {
			return commentForProcure;
		}

		public void setCommentForProcure(String commentForProcure) {
			this.commentForProcure = commentForProcure;
		}

		public String getCommentProces() {
			return commentProces;
		}

		public void setCommentProces(String commentProces) {
			this.commentProces = commentProces;
		}

		public String getHardCodes() {
			return hardCodes;
		}

		public void setHardCodes(String hardCodes) {
			this.hardCodes = hardCodes;
		}

		public String getRscs() {
			return rscs;
		}

		public void setRscs(String rscs) {
			this.rscs = rscs;
		}

		public String getState() {
			return state;
		}

		public void setState(String state) {
			this.state = state;
		}

		public String getTaxTransformCount() {
			return taxTransformCount;
		}

		public void setTaxTransformCount(String taxTransformCount) {
			this.taxTransformCount = taxTransformCount;
		}

		public String getTtaCount() {
			return ttaCount;
		}

		public void setTtaCount(String ttaCount) {
			this.ttaCount = ttaCount;
		}

		public String getComplexityProjection() {
			return complexityProjection;
		}

		public void setComplexityProjection(String complexityProjection) {
			this.complexityProjection = complexityProjection;
		}

		public String getComment() {
			return comment;
		}

		public void setComment(String comment) {
			this.comment = comment;
		}

		public String getDateSubmitJira() {
			return dateSubmitJira;
		}

		public void setDateSubmitJira(String dateSubmitJira) {
			this.dateSubmitJira = dateSubmitJira;
		}

		public String getMigrationComplDate() {
			return migrationComplDate;
		}

		public void setMigrationComplDate(String migrationComplDate) {
			this.migrationComplDate = migrationComplDate;
		}

		public String getFileDeficiences() {
			return fileDeficiences;
		}

		public void setFileDeficiences(String fileDeficiences) {
			this.fileDeficiences = fileDeficiences;
		}

		public String getFutureTtgEnhcNote() {
			return futureTtgEnhcNote;
		}

		public void setFutureTtgEnhcNote(String futureTtgEnhcNote) {
			this.futureTtgEnhcNote = futureTtgEnhcNote;
		}

		public String getRscPair() {
			return rscPair;
		}

		public void setRscPair(String rscPair) {
			this.rscPair = rscPair;
		}

		public String getReasonNtStage() {
			return reasonNtStage;
		}

		public void setReasonNtStage(String reasonNtStage) {
			this.reasonNtStage = reasonNtStage;
		}

		public String getStatus() {
			return status;
		}

		public void setStatus(String status) {
			this.status = status;
		}
	}
}
