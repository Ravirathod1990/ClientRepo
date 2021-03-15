package com.client.program;

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
import org.apache.poi.xssf.usermodel.XSSFCell;

public class BulkReadDatabaseAndUpdateExcel2 {

	private Map<String, Integer> columnMap = new HashMap<String, Integer>();
	private Map<String, RowData> rowDataMap = new ConcurrentHashMap<String, RowData>();

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

		String column1 = "FIPS";
		String column2 = "5_Document Number";
		String column3 = "DOCYR";
		String result = "FILE NAME";
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

				int idxForColumn3 = columnMap.get(column3);
				Cell column3Cell = dataRow.getCell(idxForColumn3);
				String column3Data = checkCellTypeAndReturn(column3Cell);

				RowData rowData = new RowData();
				rowData.setColumn1(column1Data);
				rowData.setColumn2(column2Data);
				rowData.setColumn3(column3Data);

				rowDataMap.put(column1Data + "_" + column2Data + "_" + column3Data, rowData);
			} else {
				physicalNumberOfRows++;
			}
		}

		ExecutorService service = Executors.newFixedThreadPool(100);
		for (Entry<String, RowData> entry : rowDataMap.entrySet()) {
			service.execute(new ProcessData(entry.getValue(), conn));
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

				int idxForColumn3 = columnMap.get(column3);
				Cell column3Cell = dataRow.getCell(idxForColumn3);
				String column3Data = checkCellTypeAndReturn(column3Cell);

				RowData rowData = rowDataMap.get(column1Data + "_" + column2Data + "_" + column3Data);

				XSSFCell cell = (XSSFCell) dataRow
						.createCell(columnMap.get(result));
				cell.setCellType(CellType.STRING);
				cell.setCellValue(rowData.getFileName());
			} else {
				physicalNumberOfRows++;
			}
		}

		FileOutputStream fileOut = new FileOutputStream(new File(filePath));
		workbook.write(fileOut);
		workbook.close();
		fileOut.close();
	}

	public class ProcessData implements Runnable {
		RowData rowData;
		Connection conn;

		public ProcessData(RowData rowData, Connection conn) {
			this.rowData = rowData;
			this.conn = conn;
		}

		@Override
		public void run() {
			String fileName = getDataFromDB(conn, rowData.getColumn1(), rowData.getColumn2(), rowData.getColumn3());
			rowData.setFileName(fileName);
			rowDataMap.put(rowData.getColumn1() + "_" + rowData.getColumn2() + "_" + rowData.getColumn3(), rowData);
		}
	}

	public String getDataFromDB(Connection conn, String column1, String column2, String column3) {
		Statement statement = null;
		try {
			statement = conn.createStatement();
			ResultSet resultSet = statement.executeQuery(
					"select CONCAT(a.Path, a.FileName) as FILENAME from fileindex a where a.fileid = '" + column1 + "'"); // Replace actual query here
			if (resultSet.next()) {
				System.out.println("data found");
				return resultSet.getString(1);
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
		return null;
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

		String filePath = "E:\\client\\files\\new\\new-12-03-21.xlsx"; // Replace actual file here
		String sheetName = "sheet1";
		
		BulkReadDatabaseAndUpdateExcel2 dataClass = new BulkReadDatabaseAndUpdateExcel2();
		dataClass.readDbDataAndMatchAndCopyToExcel(filePath, sheetName);

		System.out.println("Process completed successfully");
	}

	// Replace your database and connection url in below code
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

	public class RowData {
		private String column1 = StringUtils.EMPTY;
		private String column2 = StringUtils.EMPTY;
		private String column3 = StringUtils.EMPTY;
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
	}
}
