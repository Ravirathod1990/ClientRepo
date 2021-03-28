package com.client.program;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class BulkReadDatabaseAndUpdateExcel3 {

	static final Map<String, List<String>> params = new HashMap<>();
	private Map<String, Integer> columnMap = new HashMap<String, Integer>();
	private Map<String, RowData> rowDataMap = new ConcurrentHashMap<String, RowData>();

	public void readDbDataAndMatchAndCopyToExcel(String dbserver, String dbname, String username, String pwd, String excelpath, String sheetName, String fipscolumn, String yearcolumn, String resultcolumn, String docunumcolumn,
			String bookcolumn, String pagecolumn) throws Exception {
		processData(dbserver, dbname, username, pwd, excelpath, sheetName, fipscolumn, yearcolumn, resultcolumn, docunumcolumn, bookcolumn, pagecolumn);
	}

	public void readDbDataAndMatchAndCopyToExcel(String dbserver, String dbname, String username, String pwd, String excelpath, String sheetName, String fipscolumn, String yearcolumn, String resultcolumn, String docunumcolumn)
			throws Exception {
		processData(dbserver, dbname, username, pwd, excelpath, sheetName, fipscolumn, yearcolumn, resultcolumn, docunumcolumn, null, null);
	}

	public void readDbDataAndMatchAndCopyToExcel(String dbserver, String dbname, String username, String pwd, String excelpath, String sheetName, String fipscolumn, String yearcolumn, String resultcolumn, String bookcolumn,
			String pagecolumn) throws Exception {
		processData(dbserver, dbname, username, pwd, excelpath, sheetName, fipscolumn, yearcolumn, resultcolumn, null, bookcolumn, pagecolumn);
	}

	private void processData(String dbserver, String dbname, String username, String pwd, String excelpath, String sheetName, String fipscolumn, String yearcolumn, String resultcolumn, String docunumcolumn, String bookcolumn,
			String pagecolumn) throws IOException, FileNotFoundException {
		Connection conn = createDBConnection(dbserver, dbname, username, pwd);

		Workbook workbook = WorkbookFactory.create(new FileInputStream(excelpath));
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			if (cell != null) {
				columnMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
			}
		}

		Cell fipsCell = row.getCell(CellReference.convertColStringToIndex(fipscolumn));
		columnMap.put(fipsCell.getStringCellValue().trim(), fipsCell.getColumnIndex());
		String column1 = fipsCell.getStringCellValue().trim();

		Cell yearCell = row.getCell(CellReference.convertColStringToIndex(yearcolumn));
		columnMap.put(yearCell.getStringCellValue().trim(), yearCell.getColumnIndex());
		String column2 = yearCell.getStringCellValue().trim();

		boolean q1 = true;
		boolean q2 = true;

		String column3 = null;
		if (docunumcolumn != null) {
			Cell docnumCell = row.getCell(CellReference.convertColStringToIndex(docunumcolumn));
			columnMap.put(docnumCell.getStringCellValue().trim(), docnumCell.getColumnIndex());
			column3 = docnumCell.getStringCellValue().trim();
		} else {
			q1 = false;
		}

		String column4 = null;
		String column5 = null;
		if (bookcolumn != null && pagecolumn != null) {
			Cell bookCell = row.getCell(CellReference.convertColStringToIndex(bookcolumn));
			columnMap.put(bookCell.getStringCellValue().trim(), bookCell.getColumnIndex());
			column4 = bookCell.getStringCellValue().trim();

			Cell pageCell = row.getCell(CellReference.convertColStringToIndex(pagecolumn));
			columnMap.put(pageCell.getStringCellValue().trim(), pageCell.getColumnIndex());
			column5 = pageCell.getStringCellValue().trim();
		} else {
			q2 = false;
		}

		Cell resultCell = row.getCell(CellReference.convertColStringToIndex(resultcolumn));
		columnMap.put(resultCell.getStringCellValue().trim(), resultCell.getColumnIndex());
		String result = resultCell.getStringCellValue().trim();

		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForColumn1 = columnMap.get(column1);
				Cell column1Cell = dataRow.getCell(idxForColumn1);
				String column1Data = checkCellTypeAndReturn(column1Cell);

				int idxForColumn2 = columnMap.get(column2);
				Cell column2Cell = dataRow.getCell(idxForColumn2);
				String column2Data = checkCellTypeAndReturn(column2Cell);

				String column3Data = StringUtils.EMPTY;
				if (column3 != null) {
					int idxForColumn3 = columnMap.get(column3);
					Cell column3Cell = dataRow.getCell(idxForColumn3);
					column3Data = checkCellTypeAndReturn(column3Cell);
				}

				String column4Data = StringUtils.EMPTY;
				if (column4 != null) {
					int idxForColumn4 = columnMap.get(column4);
					Cell column4Cell = dataRow.getCell(idxForColumn4);
					column4Data = checkCellTypeAndReturn(column4Cell);
				}

				String column5Data = StringUtils.EMPTY;
				if (column5 != null) {
					int idxForColumn5 = columnMap.get(column5);
					Cell column5Cell = dataRow.getCell(idxForColumn5);
					column5Data = checkCellTypeAndReturn(column5Cell);
				}

				RowData rowData = new RowData();
				rowData.setColumn1(column1Data);
				rowData.setColumn2(column2Data);
				rowData.setColumn3(column3Data);
				rowData.setColumn4(column4Data);
				rowData.setColumn5(column5Data);

				rowDataMap.put(column1Data + "_" + column2Data + "_" + column3Data + "_" + column4Data + "_" + column5Data, rowData);
			} else {
				physicalNumberOfRows++;
			}
		}

		ExecutorService service = Executors.newFixedThreadPool(100);
		for (Entry<String, RowData> entry : rowDataMap.entrySet()) {
			service.execute(new ProcessData(entry.getValue(), conn, q1, q2));
		}

		service.shutdown();
		try {
			service.awaitTermination(Long.MAX_VALUE, TimeUnit.MILLISECONDS);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForColumn1 = columnMap.get(column1);
				Cell column1Cell = dataRow.getCell(idxForColumn1);
				String column1Data = checkCellTypeAndReturn(column1Cell);

				int idxForColumn2 = columnMap.get(column2);
				Cell column2Cell = dataRow.getCell(idxForColumn2);
				String column2Data = checkCellTypeAndReturn(column2Cell);

				String column3Data = StringUtils.EMPTY;
				if (column3 != null) {
					int idxForColumn3 = columnMap.get(column3);
					Cell column3Cell = dataRow.getCell(idxForColumn3);
					column3Data = checkCellTypeAndReturn(column3Cell);
				}

				String column4Data = StringUtils.EMPTY;
				if (column4 != null) {
					int idxForColumn4 = columnMap.get(column4);
					Cell column4Cell = dataRow.getCell(idxForColumn4);
					column4Data = checkCellTypeAndReturn(column4Cell);
				}

				String column5Data = StringUtils.EMPTY;
				if (column5 != null) {
					int idxForColumn5 = columnMap.get(column5);
					Cell column5Cell = dataRow.getCell(idxForColumn5);
					column5Data = checkCellTypeAndReturn(column5Cell);
				}
				RowData rowData = rowDataMap.get(column1Data + "_" + column2Data + "_" + column3Data + "_" + column4Data + "_" + column5Data);

				XSSFCell cell = (XSSFCell) dataRow
						.createCell(columnMap.get(result));
				cell.setCellType(CellType.STRING);
				cell.setCellValue(rowData.getFileName());
			} else {
				physicalNumberOfRows++;
			}
		}

		FileOutputStream fileOut = new FileOutputStream(new File(excelpath));
		workbook.write(fileOut);
		workbook.close();
		fileOut.close();
	}

	public class ProcessData implements Runnable {
		RowData rowData;
		Connection conn;
		boolean q1;
		boolean q2;

		public ProcessData(RowData rowData, Connection conn, boolean q1, boolean q2) {
			this.rowData = rowData;
			this.conn = conn;
			this.q1 = q1;
			this.q2 = q2;
		}

		@Override
		public void run() {
			if (!rowData.getColumn1().equals(StringUtils.EMPTY) &&
					!rowData.getColumn2().equals(StringUtils.EMPTY) &&
					!rowData.getColumn3().equals(StringUtils.EMPTY)) {
				String fileName = getDataFromDB(conn, rowData.getColumn1(), rowData.getColumn2(), rowData.getColumn3(), rowData.getColumn4(), rowData.getColumn5(), q1, q2);
				rowData.setFileName(fileName);
				rowDataMap.put(rowData.getColumn1() + "_" + rowData.getColumn2() + "_" + rowData.getColumn3() + "_" + rowData.getColumn4() + "_" + rowData.getColumn5(), rowData);
			}
		}
	}

	public String getDataFromDB(Connection conn, String column1, String column2, String column3, String column4, String column5, boolean q1, boolean q2) {
		Statement statement = null;
		try {
			statement = conn.createStatement();
			if (q1) {
				if (!column3.equals(StringUtils.EMPTY)) {
					ResultSet resultSet = statement.executeQuery(
							"select CONCAT(a.Path, a.FileName) as FILENAME from fileindex a where a.fileid = '" + column1 + "'");
					if (resultSet.next()) {
						System.out.println("data found in first query");
						return resultSet.getString(1);
					} else {
						System.out.println("data not found in first query");
					}
				}
			}
			if (q2) {
				if (!column4.equals(StringUtils.EMPTY) || !column5.equals(StringUtils.EMPTY)) {
					ResultSet resultSet1 = statement.executeQuery(
							"select CONCAT(a.Path, a.FileName) as FILENAME from fileindex a where a.fileid = '" + column4 + "'");
					if (resultSet1.next()) {
						System.out.println("data found in second query");
						return resultSet1.getString(1);
					} else {
						System.out.println("data not found in second query");
					}
				}
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
		return StringUtils.EMPTY;
	}

	public String checkCellTypeAndReturn(Cell cell) {
		String cellVal = StringUtils.EMPTY;
		if (cell != null) {
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
		}
		return cellVal;
	}

	public static void main(String[] args) throws Exception {

		createArgList(args);
		String dbserver = params.get("dbserver").get(0);
		String dbname = params.get("dbname").get(0);
		String username = params.get("username").get(0);
		String pwd = params.get("pwd").get(0);
		String excelpath = params.get("excelpath").get(0);
		String sheet = params.get("sheet").get(0);
		String fipscolumn = params.get("fipscolumn").get(0);
		String yearcolumn = params.get("yearcolumn").get(0);
		String resultcolumn = params.get("resultcolumn").get(0);

		String docunumcolumn = null;
		if (params.get("docunumcolumn") != null) {
			docunumcolumn = params.get("docunumcolumn").get(0);
		}
		String bookcolumn = null;
		if (params.get("bookcolumn") != null) {
			bookcolumn = params.get("bookcolumn").get(0);
		}
		String pagecolumn = null;
		if (params.get("pagecolumn") != null) {
			pagecolumn = params.get("pagecolumn").get(0);
		}

		BulkReadDatabaseAndUpdateExcel3 dataClass = new BulkReadDatabaseAndUpdateExcel3();
		dataClass.readDbDataAndMatchAndCopyToExcel(dbserver, dbname, username, pwd, excelpath, sheet, fipscolumn, yearcolumn, resultcolumn, docunumcolumn, bookcolumn, pagecolumn);

		// dataClass.readDbDataAndMatchAndCopyToExcel(dbserver, dbname, username, pwd, excelpath, sheet, fipscolumn, yearcolumn, resultcolumn, docunumcolumn);

		// dataClass.readDbDataAndMatchAndCopyToExcel(dbserver, dbname, username, pwd, excelpath, sheet, fipscolumn, yearcolumn, resultcolumn, bookcolumn, pagecolumn);

		System.out.println("Process completed successfully");
	}

	// Replace your database and connection url in below code
	public Connection createDBConnection(String dbserver, String dbname, String username, String pwd) {
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

	public static void createArgList(String[] args) {
		List<String> options = null;
		for (int i = 0; i < args.length; i++) {
			final String a = args[i];

			if (a.startsWith("--")) {
				options = new ArrayList<>();
				params.put(a.substring(2), options);
			} else if (options != null) {
				options.add(a);
			}
		}
	}

	public class RowData {
		private String column1 = StringUtils.EMPTY;
		private String column2 = StringUtils.EMPTY;
		private String column3 = StringUtils.EMPTY;
		private String column4 = StringUtils.EMPTY;
		private String column5 = StringUtils.EMPTY;
		private String fileName = StringUtils.EMPTY;

		public String getColumn1() {
			return column1;
		}

		public void setColumn1(String column1) {
			this.column1 = column1;
		}

		public String getColumn2() {
			return column2;
		}

		public void setColumn2(String column2) {
			this.column2 = column2;
		}

		public String getColumn3() {
			return column3;
		}

		public void setColumn3(String column3) {
			this.column3 = column3;
		}

		public String getFileName() {
			return fileName;
		}

		public void setFileName(String fileName) {
			this.fileName = fileName;
		}

		// New code till last line
		public String getColumn4() {
			return column4;
		}

		public void setColumn4(String column4) {
			this.column4 = column4;
		}

		public String getColumn5() {
			return column5;
		}

		public void setColumn5(String column5) {
			this.column5 = column5;
		}
	}
}
