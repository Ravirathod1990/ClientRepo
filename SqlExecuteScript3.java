package com.client.program;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

public class SqlExecuteScript3 {

	public static void main(String[] args) {

		String userName = "admin";
		String password = "admin";
		String path = "C:\\Program Files\\Microsoft SQL Server\\MSSQL15.SQLEXPRESS\\MSSQL\\Backup\\fold";
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
					segregateCreateInsert(conn, path);
				}

				if (isDirectory) {
					File folder = new File(path);
					File[] listOfFiles = folder.listFiles();

					for (int i = 0; i < listOfFiles.length; i++) {
						segregateCreateInsert(conn, path + "\\" + listOfFiles[i].getName());
					}
				}
				System.out.println("Script Executed");
			}
		} catch (SQLException | IOException ex) {
			System.out.println("An error occurred.");
			ex.printStackTrace();
		}
	}

	private static void segregateCreateInsert(Connection conn, String fileName) throws IOException, SQLException {
		BufferedReader in = new BufferedReader(new FileReader(fileName));
		String str;
		boolean isCreate = false;
		Path tempCreateFile = Files.createTempFile("schema-" + UUID.randomUUID(), ".tmp");
		Path tempInsertFile = Files.createTempFile("insert-" + UUID.randomUUID(), ".tmp");

		while ((str = in.readLine()) != null) {
			if (str.startsWith("CREATE TABLE")) {
				isCreate = true;
			}
			if (isCreate) {
				Files.write(tempCreateFile, str.getBytes(StandardCharsets.UTF_8), StandardOpenOption.APPEND);
				Files.write(tempCreateFile, System.lineSeparator().getBytes(StandardCharsets.UTF_8), StandardOpenOption.APPEND);
			}
			if (!isCreate && (!str.equals(");") || !str.equals(")"))) {
				Files.write(tempInsertFile, str.getBytes(StandardCharsets.UTF_8), StandardOpenOption.APPEND);
				Files.write(tempInsertFile, System.lineSeparator().getBytes(StandardCharsets.UTF_8), StandardOpenOption.APPEND);
			}
			if (isCreate && (str.equals(");") || str.equals(")"))) {
				isCreate = false;
			}
		}
		in.close();

		createTableUsingScript(conn, tempCreateFile);
		insertDataUsingScript(conn, tempInsertFile);
	}

	private static void insertDataUsingScript(Connection conn, Path fileName) throws FileNotFoundException, IOException, SQLException {
		BufferedReader in = new BufferedReader(new FileReader(fileName.toString()));
		String str;
		StringBuffer sb = new StringBuffer();
		List<String> recordList = new ArrayList<String>();
		int i = 0;

		while ((str = in.readLine()) != null) {
			if (i != 0) {
				if (str.startsWith("INSERT INTO")) {
					recordList.add(sb.toString());
					sb = new StringBuffer();
				} else {
					sb.append("\n ");
				}
			}
			sb.append(str);
			i++;
		}
		recordList.add(sb.toString());
		in.close();
		Statement stmt = conn.createStatement();
		for (String record : recordList) {
			stmt.execute(record);
		}
		stmt.closeOnCompletion();
		fileName.toFile().delete();
	}

	private static void createTableUsingScript(Connection conn, Path fileName) throws FileNotFoundException, IOException, SQLException {
		BufferedReader in = new BufferedReader(new FileReader(fileName.toString()));
		String str;
		StringBuffer sb = new StringBuffer();
		while ((str = in.readLine()) != null) {
			sb.append(str + "\n ");
		}
		in.close();
		Statement stmt = conn.createStatement();
		stmt.execute(sb.toString());
		stmt.closeOnCompletion();
		fileName.toFile().delete();
	}
}
