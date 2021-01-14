package com.client.program;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

public class SqlDataExtractorAndFileGenerator4Derby {

	public static void main(String[] args) {

		String userName = "admin"; // Pass valid username
		String password = "admin"; // Pass valid password
		String destFile = "C:\\Program Files\\Microsoft SQL Server\\MSSQL15.SQLEXPRESS\\MSSQL\\Backup\\backup.sql"; // Pass valid destination file location
		String databaseName = "jpademo"; // Pass valid database name
		String url = "jdbc:sqlserver://localhost\\SQLEXPRESS:1433;database=" + databaseName + ";integratedSecurity=true"; // Pass valid connection string

		extractDataAndGenerateSql(userName, password, destFile, databaseName, url);
	}

	private static void extractDataAndGenerateSql(String userName, String password, String destFile, String databaseName, String url) {
		Connection conn = null;
		try {
			conn = DriverManager.getConnection(url, userName, password);
			if (conn != null) {
				System.out.println("Connected to the database");

				DatabaseMetaData md = conn.getMetaData();
				ResultSet rs = md.getTables(databaseName, "dbo", "%", new String[] { "TABLE" });
				StringBuilder sb = new StringBuilder();

				List<String> tableNameList = new ArrayList<>();
				while (rs.next()) {
					tableNameList.add(rs.getString(3));
				}
				for (int j = 0; j < tableNameList.size(); j++) {
					String tableName = tableNameList.get(j);
					Statement st = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
					String sql = "select * from " + tableName;
					ResultSet rsSub = st.executeQuery(sql);
					ResultSetMetaData metaData = rsSub.getMetaData();
					int columnCount = metaData.getColumnCount();

					sb.append("CREATE TABLE " + tableName + "(");
					sb.append(System.lineSeparator());
					for (int i = 1; i <= columnCount; i++) {
						String dataType = getDataType(metaData.getColumnTypeName(i));
						if (metaData.getColumnName(i).equals("year") || metaData.getColumnName(i).equals("date") || metaData.getColumnName(i).equals("day")) {
							sb.append("\t" + '"' + metaData.getColumnName(i) + '"' + " " + dataType);
						} else {
							sb.append("\t" + metaData.getColumnName(i) + " " + dataType);
						}
						if (dataType.equals("VARCHAR") && metaData.getColumnDisplaySize(i) > 255) {
							sb.append("(255)");
						} else {
							if (!(dataType.equals("INT") || dataType.equals("TIMESTAMP"))) {
								sb.append("(" + metaData.getColumnDisplaySize(i) + ")");
							}
						}
						if (i != columnCount) {
							sb.append(",");
						}
						sb.append(System.lineSeparator());
					}
					if (j == tableNameList.size() - 1) {
						sb.append(")");
					} else {
						sb.append("),");
					}
					sb.append(System.lineSeparator());

					rsSub.close();
					st.close();
					sb.append(System.lineSeparator());
				}
				writeToFile(sb, destFile);
				rs.close();
				System.out.println("File created successfully !!");
			}
		} catch (SQLException ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		}
	}

	public static String getDataType(String sqlDataType) {
		if (sqlDataType.toUpperCase().equals("NCHAR") || sqlDataType.toUpperCase().equals("NVARCHAR")) {
			return "VARCHAR";
		} else if (sqlDataType.toUpperCase().equals("BIT")) {
			return "SMALLINT";
		} else if (sqlDataType.toUpperCase().equals("DATETIME2") || sqlDataType.toUpperCase().equals("DATETIME")) {
			return "TIMESTAMP";
		}
		return sqlDataType.toUpperCase();
	}

	private static void writeToFile(StringBuilder sb, String path) {
		BufferedWriter out = null;
		try {
			File file = new File(path);
			out = new BufferedWriter(new FileWriter(file, true));
			out.write(sb.toString());
			out.close();
		} catch (IOException e) {
			System.out.println("writeToFile ==>> An error occurred.");
			e.printStackTrace();
		}
	}
}
