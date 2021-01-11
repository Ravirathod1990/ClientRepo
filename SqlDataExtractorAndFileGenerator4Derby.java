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

				while (rs.next()) {
					String tableName = rs.getString(3);
					Statement st = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
					String sql = "select * from " + tableName;
					ResultSet rsSub = st.executeQuery(sql);
					ResultSetMetaData metaData = rsSub.getMetaData();
					int columnCount = metaData.getColumnCount();

					ResultSet rsPrim = md.getPrimaryKeys(databaseName, "dbo", tableName);
					boolean isPrimaryExist = false;
					String pkColumnName = "";
					while (rsPrim.next()) {
						isPrimaryExist = true;
						pkColumnName = rsPrim.getString("COLUMN_NAME");
					}

					sb.append("CREATE TABLE " + tableName + "(");
					sb.append(System.lineSeparator());
					for (int i = 1; i <= columnCount; i++) {
						int nullable = metaData.isNullable(i);
						String colDef = "";
						if (nullable == ResultSetMetaData.columnNullable) {
							colDef = "NULL";
						} else if (nullable == ResultSetMetaData.columnNoNulls) {
							colDef = "NOT NULL";
						}

						String pkIdentity = "";
						if (isPrimaryExist) {
							if (metaData.getColumnName(i).equals(pkColumnName)) {
								pkIdentity = " GENERATED ALWAYS AS IDENTITY";
								sb.append("\t" + metaData.getColumnName(i) + " " + getDataType(metaData.getColumnTypeName(i)) + "(" + metaData.getColumnDisplaySize(i) + ") " + colDef + pkIdentity);
							} else {
								sb.append("\t" + metaData.getColumnName(i) + " " + getDataType(metaData.getColumnTypeName(i)) + "(" + metaData.getColumnDisplaySize(i) + ") " + colDef);
							}
						} else {
							sb.append("\t" + metaData.getColumnName(i) + " " + getDataType(metaData.getColumnTypeName(i)) + "(" + metaData.getColumnDisplaySize(i) + ") " + colDef);
						}
						if (!isPrimaryExist) {
							if (i != columnCount) {
								sb.append(",");
							}
						} else {
							sb.append(",");
						}
						sb.append(System.lineSeparator());
					}
					if (isPrimaryExist) {
						sb.append("\tPRIMARY KEY (" + pkColumnName + ")");
						sb.append(System.lineSeparator());
					}
					sb.append(")");
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
