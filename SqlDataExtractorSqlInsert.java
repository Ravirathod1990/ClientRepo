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

public class SqlDataExtractorSqlInsert {

	public static void main(String[] args) {

		String userName = "admin"; // Pass valid username
		String password = "admin"; // Pass valid password
		String destFile = "C:\\Program Files\\Microsoft SQL Server\\MSSQL15.SQLEXPRESS\\MSSQL\\Backup\\"; // Pass valid destination file location
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
				while (rs.next()) {
					StringBuilder sb = new StringBuilder();
					String tableName = rs.getString(3);
					String destTableFile = destFile + tableName + ".sql";
					Statement st = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
					String sql = "select * from " + tableName;
					ResultSet rsSub = st.executeQuery(sql);
					ResultSetMetaData metaData = rsSub.getMetaData();
					int rowCount = metaData.getColumnCount();

					while (rsSub.next()) {
						sb.append("INSERT INTO " + tableName + "(");
						for (int i = 1; i <= rowCount; i++) {
							if (i == rowCount) {
								sb.append(metaData.getColumnName(i));
							} else {
								sb.append(metaData.getColumnName(i) + ",");
							}
						}
						sb.append(") VALUES");
						sb.append("(");
						for (int i = 1; i <= rowCount; i++) {
							String type = metaData.getColumnTypeName(i);
							if (isPrimitive(type)) {
								if (i == rowCount) {
									if (rsSub.getString(i) != null) {
										sb.append(rsSub.getString(i).trim());
									} else {
										sb.append("" + null);
									}
								} else {
									if (rsSub.getString(i) != null) {
										sb.append(rsSub.getString(i).trim() + ",");
									} else {
										sb.append(null + ",");
									}
								}
							} else {
								if (i == rowCount) {
									if (rsSub.getString(i) != null) {
										sb.append("'" + rsSub.getString(i).trim().replaceAll("'", "''") + "'");
									} else {
										sb.append("" + null);
									}
								} else {
									if (rsSub.getString(i) != null) {
										sb.append("'" + rsSub.getString(i).trim().replaceAll("'", "''") + "',");
									} else {
										sb.append(null + ",");
									}
								}
							}
						}
						sb.append(");");
						sb.append(System.lineSeparator());
					}
					rsSub.close();
					st.close();
					sb.append(System.lineSeparator());
					writeToFile(sb, destTableFile);
				}
				rs.close();
				System.out.println("File created successfully !!");
			}
		} catch (SQLException ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		}
	}

	private static boolean isPrimitive(String type) {
		return type.equals("bigint")
				|| type.equals("int")
				|| type.equals("tinyint")
				|| type.equals("smallint")
				|| type.equals("float")
				|| type.equals("decimal")
				|| type.equals("bit");
	}

	private static void writeToFile(StringBuilder sb, String path) {
		BufferedWriter out = null;
		try {
			File file = new File(path);
			out = new BufferedWriter(new FileWriter(file, false));
			out.write(sb.toString());
			out.close();
		} catch (IOException e) {
			System.out.println("writeToFile ==>> An error occurred.");
			e.printStackTrace();
		}
	}
}
