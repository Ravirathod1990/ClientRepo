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
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadAndValidateExcelWithDatabase {

	private Map<String, ViperRow> dataList = new LinkedHashMap<String, ViperRow>();

	public void readExcel(String filePath, String sheetName) throws Exception {
		Connection conn = createDBConnection();

		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));
		Sheet sheet = workbook.getSheet(sheetName);

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
			ViperRow viperRow = new ViperRow();
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForTt = map.get("TT");
				int idxForTitle = map.get("Title");
				int idxForVendor = map.get("Vendor");
				int idxForInProduction = map.get("In Production");
				int idxForEnabledDisabled = map.get("Enabled/Disabled");
				int idxForFileBillYearTyep = map.get("File Bill Year Type");
				int idxForFilePaymentStatusType = map.get("File Payment Status Type");
				int idxForFileBillType = map.get("File Bill Type");
				int idxForTimingConstraints = map.get("Timing Constrains");
				int idxForCommentForProcure = map.get("Comment for Procure");
				int idxForCommentProcesHardCode = map.get("Comment ProcessorHard codes");
				int idxForState = map.get("State(s)");
				int idxForTaxTransformCount = map.get("Tax Transform Count");
				int idxForTtaCount = map.get("TTA Count");
				int idxForComplexityProjection = map.get("Complexity Projection");
				int idxForComment = map.get("Comment");
				int idxForDateSubmitJira = map.get("Date Submitted JIRA");
				int idxForMigrationComplDate = map.get("Migration Completion Date");
				int idxForFileDeficiences = map.get("File Deficiences");
				int idxForFutureTtgEnhcNote = map.get("Future TTG Enhancement Needs");
				int idxForRscPair = map.get("RSC pair");
				int idxForReasonNtStage = map.get("Reason not Staged");
				int idxForHardCodes = map.get("Hard Codes");
				int idxForRscs = map.get("RSC(s)");

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
				Cell commentProcesHardCodeCell = dataRow.getCell(idxForCommentProcesHardCode);
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
				if (commentProcesHardCodeCell != null) {
					viperRow.setCommentProcesHardCode(checkCellTypeAndReturn(commentProcesHardCodeCell));
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
				if (hardCodesCell != null) {
					viperRow.setHardCodes(checkCellTypeAndReturn(hardCodesCell));
				}
				if (rscsCell != null) {
					viperRow.setRscs(checkCellTypeAndReturn(rscsCell));
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
			ResultSet resultSet = statement
					.executeQuery("select title,tax_transformation,vendor,file_bill_year_type,file_payment_status_type,"
							+ "file_bill_type,comment_for_procure,in_production,enabled_disabled,timing_constraints,"
							+ "comment_for_processor,hard_codes,rsc,states,tax_transform_count,tta_count,"
							+ "complexity_projection,comment,date_submitted_to_jira,migration_completion_date,"
							+ "file_deficiencies,future_ttg_enhancement_needs,rsc_pair,reason_not_staged from vipertable where title = '"
							+ title + "'");
			if (resultSet.next()) {
				System.out.println("data found");
				boolean passed = true;
				StringBuffer sbStatus = new StringBuffer();
				viperRow.setDbTitle(resultSet.getString(1));
				viperRow.setDbTt(resultSet.getString(2));
				if (!viperRow.getTt().equals(viperRow.getDbTt())) {
					sbStatus.append("TT not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbVendor(resultSet.getString(3));
				if (!viperRow.getVendor().equals(viperRow.getDbVendor())) {
					sbStatus.append("Vendor not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFileBillYearTyep(resultSet.getString(4));
				if (!viperRow.getFileBillYearTyep().equals(viperRow.getDbFileBillYearTyep())) {
					sbStatus.append("File Bill Year Type not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFilePaymentStatusType(resultSet.getString(5));
				if (!viperRow.getFilePaymentStatusType().equals(viperRow.getDbFilePaymentStatusType())) {
					sbStatus.append("File Payment Status Type not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFileBillType(resultSet.getString(6));
				if (!viperRow.getFileBillType().equals(viperRow.getDbFileBillType())) {
					sbStatus.append("File Bill Type not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbCommentForProcure(resultSet.getString(7));
				if (!viperRow.getCommentForProcure().equals(viperRow.getDbCommentForProcure())) {
					sbStatus.append("Comment for Procure not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbInProduction(resultSet.getString(8));
				if (!viperRow.getInProduction().equals(viperRow.getDbInProduction())) {
					sbStatus.append("In Production not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbEnabledDisabled(resultSet.getString(9));
				if (!viperRow.getEnabledDisabled().equals(viperRow.getDbEnabledDisabled())) {
					sbStatus.append("Enabled/Disabled not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbTimingConstraints(resultSet.getString(10));
				if (!viperRow.getTimingConstraints().equals(viperRow.getDbTimingConstraints())) {
					sbStatus.append("Timing Constrains not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbCommentForProcure(resultSet.getString(11));
				if (!viperRow.getCommentForProcure().equals(viperRow.getDbCommentForProcure())) {
					sbStatus.append("Comment for Procure not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbHardCodes(resultSet.getString(12));
				if (!viperRow.getHardCodes().equals(viperRow.getDbHardCodes())) {
					sbStatus.append("Comment ProcessorHard codes not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbRscs(resultSet.getString(13));
				if (!viperRow.getRscs().equals(viperRow.getDbRscs())) {
					sbStatus.append("RSC(s) not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbState(resultSet.getString(14));
				if (!viperRow.getState().equals(viperRow.getDbState())) {
					sbStatus.append("State(s) not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbTaxTransformCount(resultSet.getString(15));
				if (!viperRow.getTaxTransformCount().equals(viperRow.getDbTaxTransformCount())) {
					sbStatus.append("Tax Transform Count not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbTtaCount(resultSet.getString(16));
				if (!viperRow.getTtaCount().equals(viperRow.getDbTtaCount())) {
					sbStatus.append("TTA Count not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbComplexityProjection(resultSet.getString(17));
				if (!viperRow.getComplexityProjection().equals(viperRow.getDbComplexityProjection())) {
					sbStatus.append("Complexity Projection not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbComment(resultSet.getString(18));
				if (!viperRow.getComment().equals(viperRow.getDbComment())) {
					sbStatus.append("Comment not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbDateSubmitJira(resultSet.getString(19));
				if (!viperRow.getDateSubmitJira().equals(viperRow.getDbDateSubmitJira())) {
					sbStatus.append("Date Submitted JIRA not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbMigrationComplDate(resultSet.getString(20));
				if (!viperRow.getMigrationComplDate().equals(viperRow.getDbMigrationComplDate())) {
					sbStatus.append("Migration Completion Date not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFileDeficiences(resultSet.getString(21));
				if (!viperRow.getFileDeficiences().equals(viperRow.getDbFileDeficiences())) {
					sbStatus.append("File Deficiences not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbFutureTtgEnhcNote(resultSet.getString(22));
				if (!viperRow.getFutureTtgEnhcNote().equals(viperRow.getDbFutureTtgEnhcNote())) {
					sbStatus.append("Future TTG Enhancement Needs not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbRscPair(resultSet.getString(23));
				if (!viperRow.getRscPair().equals(viperRow.getDbRscPair())) {
					sbStatus.append("RSC pair not match with database; ");
					sbStatus.append(System.getProperty("line.separator"));
					passed = false;
				}
				viperRow.setDbReasonNtStage(resultSet.getString(24));
				if (!viperRow.getReasonNtStage().equals(viperRow.getDbReasonNtStage())) {
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

	public void copyFileSheet(String filePath, String targetFilePath, String sheetName) throws IOException {
		BufferedInputStream bis = new BufferedInputStream(new FileInputStream(filePath));
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
		int idxForTitle = 0;

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

							if (iRow == 0 && checkCellTypeAndReturn(cell).equalsIgnoreCase("Title")) {
								idxForTitle = cell.getColumnIndex();
							}
						}
					}
					if (iRow == 0) {
						fCell = lCell;

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
						list.add("DB Comment ProcessorHard codes");
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
						list.add("DB Hard Codes");
						list.add("DB RSC(s)");
						list.add("Status");

						lCell = lCell + 25;

						int counter = 0;
						for (int iCell = fCell; iCell < lCell; iCell++) {
							myCell = myRow.createCell(iCell);
							myCell.setCellType(CellType.STRING);
							myCell.setCellValue(list.get(counter));
							counter++;
						}
					} else {
						Cell titleCell = row.getCell(idxForTitle);
						String titleValue = checkCellTypeAndReturn(titleCell);

						ViperRow viperRow = dataList.get(titleValue);
						int counter = lCell;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbTt());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbTitle());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbVendor());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbInProduction());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbEnabledDisabled());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbFileBillYearTyep());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbFilePaymentStatusType());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbFileBillType());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbTimingConstraints());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbCommentForProcure());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbCommentProcesHardCode());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbState());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbTaxTransformCount());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbTtaCount());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbComplexityProjection());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbComment());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbDateSubmitJira());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbMigrationComplDate());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbFileDeficiences());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbFutureTtgEnhcNote());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbRscPair());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbReasonNtStage());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbHardCodes());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getDbRscs());
						counter++;
						myCell = myRow.createCell(counter);
						myCell.setCellType(CellType.STRING);
						myCell.setCellValue(viperRow.getStatus());
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
		String targetFilePath;

		FileSystem system = FileSystems.getDefault();
		Path original = system.getPath(filePath);
		String origFileName = original.getFileName().toString();
		String fileName = origFileName.substring(0, origFileName.lastIndexOf("."));
		String fileExt = origFileName.substring(origFileName.lastIndexOf("."), origFileName.length());
		Path targetFile = system.getPath(original.getParent() + "\\" + fileName + "_copy" + fileExt);
		targetFilePath = targetFile.toString();

		ReadAndValidateExcelWithDatabase dataClass = new ReadAndValidateExcelWithDatabase();
		dataClass.readExcel(filePath, sheetName);

		File file = new File(targetFilePath);
		file.createNewFile();
		dataClass.copyFileSheet(filePath, targetFilePath, sheetName);

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
		private String commentProcesHardCode = StringUtils.EMPTY;
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
		private String hardCodes = StringUtils.EMPTY;
		private String rscs = StringUtils.EMPTY;

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
		private String dbCommentProcesHardCode = StringUtils.EMPTY;
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
		private String dbHardCodes = StringUtils.EMPTY;
		private String dbRscs = StringUtils.EMPTY;

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

		public String getCommentProcesHardCode() {
			return commentProcesHardCode;
		}

		public void setCommentProcesHardCode(String commentProcesHardCode) {
			this.commentProcesHardCode = commentProcesHardCode;
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

		public String getDbCommentProcesHardCode() {
			return dbCommentProcesHardCode;
		}

		public void setDbCommentProcesHardCode(String dbCommentProcesHardCode) {
			this.dbCommentProcesHardCode = dbCommentProcesHardCode;
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

		public String getStatus() {
			return status;
		}

		public void setStatus(String status) {
			this.status = status;
		}
	}
}
