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

public class SqlDataExtractorAndFileGenerator2 {

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

					// new code
					sb.append("CREATE TABLE [dbo].[" + tableName + "](");
					sb.append(System.lineSeparator());
					for (int i = 1; i <= columnCount; i++) {
						int nullable = metaData.isNullable(i);
						String colDef = "";
						if (nullable == ResultSetMetaData.columnNullable) {
							colDef = "NULL";
						} else if (nullable == ResultSetMetaData.columnNoNulls) {
							colDef = "NOT NULL";
						}
						sb.append("\t[" + metaData.getColumnName(i) + "] [" + metaData.getColumnTypeName(i) + "](" + metaData.getColumnDisplaySize(i) + ") " + colDef + ",");
						sb.append(System.lineSeparator());
					}

					ResultSet rsPrim = md.getPrimaryKeys(databaseName, "dbo", tableName);
					while (rsPrim.next()) {
						sb.append("CONSTRAINT [" + rsPrim.getString("PK_NAME") + "] PRIMARY KEY CLUSTERED");
						sb.append(System.lineSeparator());
						sb.append("(");
						sb.append(System.lineSeparator());
						sb.append("[" + rsPrim.getString("COLUMN_NAME") + "] ASC");
						sb.append(System.lineSeparator());
						sb.append(")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]");
						sb.append(System.lineSeparator());
						sb.append(") ON [PRIMARY]");
						sb.append(System.lineSeparator());
						sb.append("GO");
						sb.append(System.lineSeparator());
					}
					sb.append(System.lineSeparator());

					ResultSet rsForg = md.getExportedKeys(databaseName, "dbo", tableName);
					while (rsForg.next()) {
						sb.append("ALTER TABLE [dbo].[" + rsForg.getString("FKTABLE_NAME") + "]  WITH CHECK ADD  CONSTRAINT [FK_" + rsForg.getString("FKTABLE_NAME") + "_" + rsForg.getString("PKTABLE_NAME") + "] FOREIGN KEY(["
								+ rsForg.getString("FKCOLUMN_NAME") + "])");
						sb.append(System.lineSeparator());
						sb.append("REFERENCES [dbo].[" + rsForg.getString("PKTABLE_NAME") + "] ([" + rsForg.getString("PKCOLUMN_NAME") + "])");
						sb.append(System.lineSeparator());
						sb.append("GO");
						sb.append(System.lineSeparator());
						sb.append("ALTER TABLE [dbo].[" + rsForg.getString("FKTABLE_NAME") + "] CHECK CONSTRAINT [FK_" + rsForg.getString("FKTABLE_NAME") + "_" + rsForg.getString("PKTABLE_NAME") + "]");
						sb.append(System.lineSeparator());
						sb.append("GO");
					}

					sb.append("INSERT INTO `" + tableName + "` (");
					for (int i = 1; i <= columnCount; i++) {
						if (i == columnCount) {
							sb.append("`" + metaData.getColumnName(i) + "`");
						} else {
							sb.append("`" + metaData.getColumnName(i) + "`,");
						}
					}
					sb.append(") VALUES");
					sb.append(System.lineSeparator());
					while (rsSub.next()) {
						sb.append("(");
						for (int i = 1; i <= columnCount; i++) {
							String type = metaData.getColumnTypeName(i);
							if (isPrimitive(type)) {
								if (i == columnCount) {
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
								if (i == columnCount) {
									if (rsSub.getString(i) != null) {
										sb.append("'" + rsSub.getString(i).trim() + "'");
									} else {
										sb.append("" + null);
									}
								} else {
									if (rsSub.getString(i) != null) {
										sb.append("'" + rsSub.getString(i).trim() + "',");
									} else {
										sb.append(null + ",");
									}
								}
							}
						}
						if (rsSub.isLast()) {
							sb.append(");");
						} else {
							sb.append("),");
						}
						sb.append(System.lineSeparator());
					}
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
			out = new BufferedWriter(new FileWriter(file, true));
			out.write(sb.toString());
			out.close();
		} catch (IOException e) {
			System.out.println("writeToFile ==>> An error occurred.");
			e.printStackTrace();
		}
	}
}
