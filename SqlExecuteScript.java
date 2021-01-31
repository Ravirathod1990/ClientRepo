package com.client.program;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class SqlExecuteScript {

	public static void main(String[] args) {

		String userName = "admin";
		String password = "admin";
		String path = "C:\\Program Files\\Microsoft SQL Server\\MSSQL15.SQLEXPRESS\\MSSQL\\Backup\\backup.sql";
		String databaseName = "jpademo";
		String url = "jdbc:sqlserver://localhost\\SQLEXPRESS:1433;database=" + databaseName + ";integratedSecurity=true";

		Connection conn = null;
		try {
			conn = DriverManager.getConnection(url, userName, password);
			if (conn != null) {
				System.out.println("Connected to the database");

				Path filePath = Paths.get(path);
				boolean isDirectory = Files.isDirectory(filePath);
				boolean isFile = Files.isRegularFile(filePath);

				if (isFile) {
					createTableUsingScript(conn, path);
				}

				if (isDirectory) {
					File folder = new File(path);
					File[] listOfFiles = folder.listFiles();

					for (int i = 0; i < listOfFiles.length; i++) {
						createTableUsingScript(conn, path + "\\" + listOfFiles[i].getName());
					}
				}
				System.out.println("Script Executed");
			}
		} catch (SQLException | IOException ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		}
	}

	private static void createTableUsingScript(Connection conn, String fileName) throws FileNotFoundException, IOException, SQLException {
		BufferedReader in = new BufferedReader(new FileReader(fileName));
		String str;
		StringBuffer sb = new StringBuffer();
		while ((str = in.readLine()) != null) {
			sb.append(str + "\n ");
		}
		in.close();
		Statement stmt = conn.createStatement();
		stmt.execute(sb.toString());
	}
}
