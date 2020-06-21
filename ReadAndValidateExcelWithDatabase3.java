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

public class ReadAndValidateExcelWithDatabase3 {

	private Map<String, ViperRow> dataList = new LinkedHashMap<String, ViperRow>();
	Map<String, Integer> columnMap = new HashMap<String, Integer>();

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

				if (viperRow.getTitle() != null) {
					getDbRecordCount(conn, viperRow, viperRow.getTitle());
					dataList.put(viperRow.getTitle(), viperRow);
				}
			} else {
				physicalNumberOfRows++;
			}
		}
	}

	public List<String> getDbRecordCount(Connection conn, ViperRow viperRow, String title) {
		List<String> dbRecordCountList = new ArrayList<>();

		try {
			Statement statement = conn.createStatement();
			ResultSet resultSet = statement.executeQuery(
					"select title,tax_transformation,vendor,in_production,enabled_disabled,file_bill_year_type,file_payment_status_type,"
							+ "file_bill_type,timing_constraints,comment_for_procure,"
							+ "comment_for_processor,hard_codes,rsc,states,tax_transform_count,tta_count,"
							+ "complexity_projection,comment,date_submitted_to_jira,migration_completion_date,"
							+ "file_deficience,future_ttg_enhancement_needs,rsc_pair,reason_not_staged from vipertable where title = '"
							+ title + "'");
			if (resultSet.next()) {
				System.out.println("data found");
				boolean passed = true;
				StringBuffer sbStatus = new StringBuffer();
				viperRow.setDbTitle(resultSet.getString(1));
				viperRow.setDbTt(resultSet.getString(2));
				if (!viperRow.getTt().equals(viperRow.getDbTt())) {
					viperRow.getDataNotMatchList().add("TT");
					sbStatus.append("TT not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbVendor(resultSet.getString(3));
				if (!viperRow.getVendor().equals(viperRow.getDbVendor())) {
					viperRow.getDataNotMatchList().add("Vendor");
					sbStatus.append("Vendor not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbInProduction(resultSet.getString(4));
				if (!viperRow.getInProduction().equals(viperRow.getDbInProduction())) {
					viperRow.getDataNotMatchList().add("In Production");
					sbStatus.append("In Production not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbEnabledDisabled(resultSet.getString(5));
				if (!viperRow.getEnabledDisabled().equals(viperRow.getDbEnabledDisabled())) {
					viperRow.getDataNotMatchList().add("Enabled/Disabled");
					sbStatus.append("Enabled/Disabled not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFileBillYearTyep(resultSet.getString(6));
				if (!viperRow.getFileBillYearTyep().equals(viperRow.getDbFileBillYearTyep())) {
					viperRow.getDataNotMatchList().add("File Bill Year Type");
					sbStatus.append("File Bill Year Type not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFilePaymentStatusType(resultSet.getString(7));
				if (!viperRow.getFilePaymentStatusType().equals(viperRow.getDbFilePaymentStatusType())) {
					viperRow.getDataNotMatchList().add("File Payment Status Type");
					sbStatus.append("File Payment Status Type not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFileBillType(resultSet.getString(8));
				if (!viperRow.getFileBillType().equals(viperRow.getDbFileBillType())) {
					viperRow.getDataNotMatchList().add("File Bill Type");
					sbStatus.append("File Bill Type not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbTimingConstraints(resultSet.getString(9));
				if (!viperRow.getTimingConstraints().equals(viperRow.getDbTimingConstraints())) {
					viperRow.getDataNotMatchList().add("Timing Constrains");
					sbStatus.append("Timing Constrains not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbCommentForProcure(resultSet.getString(10));
				if (!viperRow.getCommentForProcure().equals(viperRow.getDbCommentForProcure())) {
					viperRow.getDataNotMatchList().add("Comment for Procure");
					sbStatus.append("Comment for Procure not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbCommentProces(resultSet.getString(11));
				if (!viperRow.getCommentProces().equals(viperRow.getDbCommentProces())) {
					viperRow.getDataNotMatchList().add("Comment Processor");
					sbStatus.append("Comment Processor not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbHardCodes(resultSet.getString(12));
				if (!viperRow.getHardCodes().equals(viperRow.getDbHardCodes())) {
					viperRow.getDataNotMatchList().add("Hard codes");
					sbStatus.append("Hard codes not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbRscs(resultSet.getString(13));
				if (!viperRow.getRscs().equals(viperRow.getDbRscs())) {
					viperRow.getDataNotMatchList().add("RSC(s)");
					sbStatus.append("RSC(s) not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbState(resultSet.getString(14));
				if (!viperRow.getState().equals(viperRow.getDbState())) {
					viperRow.getDataNotMatchList().add("State(s)");
					sbStatus.append("State(s) not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbTaxTransformCount(resultSet.getString(15));
				if (!viperRow.getTaxTransformCount().equals(viperRow.getDbTaxTransformCount())) {
					viperRow.getDataNotMatchList().add("Tax Transform Count");
					sbStatus.append("Tax Transform Count not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbTtaCount(resultSet.getString(16));
				if (!viperRow.getTtaCount().equals(viperRow.getDbTtaCount())) {
					viperRow.getDataNotMatchList().add("TTA Count");
					sbStatus.append("TTA Count not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbComplexityProjection(resultSet.getString(17));
				if (!viperRow.getComplexityProjection().equals(viperRow.getDbComplexityProjection())) {
					viperRow.getDataNotMatchList().add("Complexity Projection");
					sbStatus.append("Complexity Projection not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbComment(resultSet.getString(18));
				if (!viperRow.getComment().equals(viperRow.getDbComment())) {
					viperRow.getDataNotMatchList().add("Comment");
					sbStatus.append("Comment not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbDateSubmitJira(resultSet.getString(19));
				if (!viperRow.getDateSubmitJira().equals(viperRow.getDbDateSubmitJira())) {
					viperRow.getDataNotMatchList().add("Date Submitted JIRA");
					sbStatus.append("Date Submitted JIRA not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbMigrationComplDate(resultSet.getString(20));
				if (!viperRow.getMigrationComplDate().equals(viperRow.getDbMigrationComplDate())) {
					viperRow.getDataNotMatchList().add("Migration Completion Date");
					sbStatus.append("Migration Completion Date not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFileDeficiences(resultSet.getString(21));
				if (!viperRow.getFileDeficiences().equals(viperRow.getDbFileDeficiences())) {
					viperRow.getDataNotMatchList().add("File Deficiences");
					sbStatus.append("File Deficiences not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFutureTtgEnhcNote(resultSet.getString(22));
				if (!viperRow.getFutureTtgEnhcNote().equals(viperRow.getDbFutureTtgEnhcNote())) {
					viperRow.getDataNotMatchList().add("Future TTG Enhancement Needs");
					sbStatus.append("Future TTG Enhancement Needs not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbRscPair(resultSet.getString(23));
				if (!viperRow.getRscPair().equals(viperRow.getDbRscPair())) {
					viperRow.getDataNotMatchList().add("RSC pair");
					sbStatus.append("RSC pair not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbReasonNtStage(resultSet.getString(24));
				if (!viperRow.getReasonNtStage().equals(viperRow.getDbReasonNtStage())) {
					viperRow.getDataNotMatchList().add("Reason not Staged");
					sbStatus.append("Reason not Staged not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				if (passed) {
					viperRow.setStatus("Passed");
				}
				if (!passed) {
					viperRow.setStatus(sbStatus.toString());
				}
			} else {
				System.out.println("data not found");
				viperRow.setStatus("Data not available in database");
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
		int idxForTitle = 0;

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

							if (iRow == 0 && checkCellTypeAndReturn(cell).equalsIgnoreCase("Title")) {
								idxForTitle = cell.getColumnIndex();
							}
						}
					}
					if (iRow == 0) {
						myCell = myRow.createCell(lCell);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue("Status");

						List<String> list = new ArrayList<>();
						list.add("DB TT");
						list.add("DB Title");
						list.add("DB Vendor");
						list.add("DB In Production");
						list.add("DB Enabled/Disabled");
						list.add("DB File Bill Year Type");
						list.add("DB File Payment Status Type");
						list.add("DB File Bill Type");
						list.add("DB Timing Constrains");
						list.add("DB Comment for Procure");
						list.add("DB Comment Processor");
						list.add("DB Hard codes");
						list.add("DB RSC(s)");
						list.add("DB State(s)");
						list.add("DB Tax Transform Count");
						list.add("DB TTA Count");
						list.add("DB Complexity Projection");
						list.add("DB Comment");
						list.add("DB Date Submitted JIRA");
						list.add("DB Migration Completion Date");
						list.add("DB File Deficiences");
						list.add("DB Future TTG Enhancement Needs");
						list.add("DB RSC pair");
						list.add("DB Reason not Staged");

						int counter = 0;
						for (int iCell = fCell; iCell < lCell; iCell++) {
							dbCell = dbRow.createCell(iCell);
							dbCell.setCellType(CellType.STRING);
							dbCell.setCellValue(list.get(counter));
							counter++;
						}
					} else {
						Cell titleCell = row.getCell(idxForTitle);
						String titleValue = checkCellTypeAndReturn(titleCell);

						ViperRow viperRow = dataList.get(titleValue);
						int counter = 0;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbTt());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbTitle());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbVendor());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbInProduction());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbEnabledDisabled());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbFileBillYearTyep());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbFilePaymentStatusType());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbFileBillType());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbTimingConstraints());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbCommentForProcure());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbCommentProces());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbHardCodes());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbRscs());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbState());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbTaxTransformCount());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbTtaCount());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbComplexityProjection());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbComment());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbDateSubmitJira());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbMigrationComplDate());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbFileDeficiences());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbFutureTtgEnhcNote());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbRscPair());
						counter++;
						dbCell = dbRow.createCell(counter);
						dbCell.setCellType(CellType.STRING);
						dbCell.setCellValue(viperRow.getDbReasonNtStage());

						myCell = myRow.createCell(lCell);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getStatus());

						for (String columnName : viperRow.getDataNotMatchList()) {
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

		String filePath = "E:\\client\\files\\newexcel\\Viper.xlsx";
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

		ReadAndValidateExcelWithDatabase3 dataClass = new ReadAndValidateExcelWithDatabase3();
		dataClass.readExcel(filePath, sheetName);

		File file = new File(targetFilePath);
		file.createNewFile();
		dataClass.copyFileSheet(filePath, targetFilePath, sheetName, dbSheetName);

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

		private String dbTt = StringUtils.EMPTY;
		private String dbTitle = StringUtils.EMPTY;
		private String dbVendor = StringUtils.EMPTY;
		private String dbInProduction = StringUtils.EMPTY;
		private String dbEnabledDisabled = StringUtils.EMPTY;
		private String dbFileBillYearTyep = StringUtils.EMPTY;
		private String dbFilePaymentStatusType = StringUtils.EMPTY;
		private String dbFileBillType = StringUtils.EMPTY;
		private String dbTimingConstraints = StringUtils.EMPTY;
		private String dbCommentForProcure = StringUtils.EMPTY;
		private String dbCommentProces = StringUtils.EMPTY;
		private String dbHardCodes = StringUtils.EMPTY;
		private String dbRscs = StringUtils.EMPTY;
		private String dbState = StringUtils.EMPTY;
		private String dbTaxTransformCount = StringUtils.EMPTY;
		private String dbTtaCount = StringUtils.EMPTY;
		private String dbComplexityProjection = StringUtils.EMPTY;
		private String dbComment = StringUtils.EMPTY;
		private String dbDateSubmitJira = StringUtils.EMPTY;
		private String dbMigrationComplDate = StringUtils.EMPTY;
		private String dbFileDeficiences = StringUtils.EMPTY;
		private String dbFutureTtgEnhcNote = StringUtils.EMPTY;
		private String dbRscPair = StringUtils.EMPTY;
		private String dbReasonNtStage = StringUtils.EMPTY;

		private String status = StringUtils.EMPTY;
		private List<String> dataNotMatchList = new ArrayList<String>();

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

		public String getDbTt() {
			return dbTt;
		}

		public void setDbTt(String dbTt) {
			this.dbTt = dbTt;
		}

		public String getDbTitle() {
			return dbTitle;
		}

		public void setDbTitle(String dbTitle) {
			this.dbTitle = dbTitle;
		}

		public String getDbVendor() {
			return dbVendor;
		}

		public void setDbVendor(String dbVendor) {
			this.dbVendor = dbVendor;
		}

		public String getDbInProduction() {
			return dbInProduction;
		}

		public void setDbInProduction(String dbInProduction) {
			this.dbInProduction = dbInProduction;
		}

		public String getDbEnabledDisabled() {
			return dbEnabledDisabled;
		}

		public void setDbEnabledDisabled(String dbEnabledDisabled) {
			this.dbEnabledDisabled = dbEnabledDisabled;
		}

		public String getDbFileBillYearTyep() {
			return dbFileBillYearTyep;
		}

		public void setDbFileBillYearTyep(String dbFileBillYearTyep) {
			this.dbFileBillYearTyep = dbFileBillYearTyep;
		}

		public String getDbFilePaymentStatusType() {
			return dbFilePaymentStatusType;
		}

		public void setDbFilePaymentStatusType(String dbFilePaymentStatusType) {
			this.dbFilePaymentStatusType = dbFilePaymentStatusType;
		}

		public String getDbFileBillType() {
			return dbFileBillType;
		}

		public void setDbFileBillType(String dbFileBillType) {
			this.dbFileBillType = dbFileBillType;
		}

		public String getDbTimingConstraints() {
			return dbTimingConstraints;
		}

		public void setDbTimingConstraints(String dbTimingConstraints) {
			this.dbTimingConstraints = dbTimingConstraints;
		}

		public String getDbCommentForProcure() {
			return dbCommentForProcure;
		}

		public void setDbCommentForProcure(String dbCommentForProcure) {
			this.dbCommentForProcure = dbCommentForProcure;
		}

		public String getDbCommentProces() {
			return dbCommentProces;
		}

		public void setDbCommentProces(String dbCommentProces) {
			this.dbCommentProces = dbCommentProces;
		}

		public String getDbHardCodes() {
			return dbHardCodes;
		}

		public void setDbHardCodes(String dbHardCodes) {
			this.dbHardCodes = dbHardCodes;
		}

		public String getDbRscs() {
			return dbRscs;
		}

		public void setDbRscs(String dbRscs) {
			this.dbRscs = dbRscs;
		}

		public String getDbState() {
			return dbState;
		}

		public void setDbState(String dbState) {
			this.dbState = dbState;
		}

		public String getDbTaxTransformCount() {
			return dbTaxTransformCount;
		}

		public void setDbTaxTransformCount(String dbTaxTransformCount) {
			this.dbTaxTransformCount = dbTaxTransformCount;
		}

		public String getDbTtaCount() {
			return dbTtaCount;
		}

		public void setDbTtaCount(String dbTtaCount) {
			this.dbTtaCount = dbTtaCount;
		}

		public String getDbComplexityProjection() {
			return dbComplexityProjection;
		}

		public void setDbComplexityProjection(String dbComplexityProjection) {
			this.dbComplexityProjection = dbComplexityProjection;
		}

		public String getDbComment() {
			return dbComment;
		}

		public void setDbComment(String dbComment) {
			this.dbComment = dbComment;
		}

		public String getDbDateSubmitJira() {
			return dbDateSubmitJira;
		}

		public void setDbDateSubmitJira(String dbDateSubmitJira) {
			this.dbDateSubmitJira = dbDateSubmitJira;
		}

		public String getDbMigrationComplDate() {
			return dbMigrationComplDate;
		}

		public void setDbMigrationComplDate(String dbMigrationComplDate) {
			this.dbMigrationComplDate = dbMigrationComplDate;
		}

		public String getDbFileDeficiences() {
			return dbFileDeficiences;
		}

		public void setDbFileDeficiences(String dbFileDeficiences) {
			this.dbFileDeficiences = dbFileDeficiences;
		}

		public String getDbFutureTtgEnhcNote() {
			return dbFutureTtgEnhcNote;
		}

		public void setDbFutureTtgEnhcNote(String dbFutureTtgEnhcNote) {
			this.dbFutureTtgEnhcNote = dbFutureTtgEnhcNote;
		}

		public String getDbRscPair() {
			return dbRscPair;
		}

		public void setDbRscPair(String dbRscPair) {
			this.dbRscPair = dbRscPair;
		}

		public String getDbReasonNtStage() {
			return dbReasonNtStage;
		}

		public void setDbReasonNtStage(String dbReasonNtStage) {
			this.dbReasonNtStage = dbReasonNtStage;
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
