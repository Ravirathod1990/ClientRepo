package com.client.program;

import java.io.File;
import java.io.IOException;
import java.sql.SQLException;
import java.util.Properties;

import com.smattme.MysqlExportService;

public class MysqlConnectAndBackup {

	// 	Add below library in POM
	//	<dependency>
	//		<groupId>com.smattme</groupId>
	//		<artifactId>mysql-backup4j</artifactId>
	//		<version>1.2.0</version>
	//	</dependency>

	public static void main(String[] args) {

		Properties properties = new Properties();
		properties.setProperty(MysqlExportService.DB_USERNAME, "{username}");

		properties.setProperty(MysqlExportService.DB_PASSWORD, "{password}");
		properties.setProperty(MysqlExportService.JDBC_CONNECTION_STRING, "jdbc:mysql://{ip or host name and port if any}/{database name}");

		properties.setProperty(MysqlExportService.TEMP_DIR, new File("{folder path for backup}").getPath());
		properties.setProperty(MysqlExportService.PRESERVE_GENERATED_ZIP, "true");
		MysqlExportService mysqlExportService = new MysqlExportService(properties);

		try {
			mysqlExportService.export();
		} catch (ClassNotFoundException | IOException | SQLException e) {
			e.printStackTrace();
		}
	}
}
