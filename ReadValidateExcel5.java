import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadValidateExcel5 {

	private List<TdoReportRow> listOfRdoDataFromReport = new ArrayList<TdoReportRow>();

	public String readAndValidateExcelData(String tdoFilepath, String tdoSheetName) throws Exception {

		if (!new File(tdoFilepath).exists()) {
			return "TDO file not exist in directory";
		}
		readRdoReportExcel(tdoFilepath, tdoSheetName);
		return "";
	}

	public void readRdoReportExcel(String filePath, String tdoSheetName) throws Exception {
		Workbook workbook = WorkbookFactory.create(new FileInputStream(filePath));

		Sheet sheet = null;
		if (tdoSheetName == null || tdoSheetName.equals("")) {
			sheet = workbook.getSheetAt(0);
		} else {
			sheet = workbook.getSheet(tdoSheetName);
		}

		Map<String, Integer> map = new HashMap<String, Integer>();
		Row row = sheet.getRow(0);
		
		if (row == null) {
			System.out.println("Data not found in sheet");
			return;
		}

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		for (int x = 1; x < sheet.getPhysicalNumberOfRows(); x++) {
			TdoReportRow reportRow = new TdoReportRow();
			Row dataRow = sheet.getRow(x);

			int idxForVersion = map.get("Version");
			int idxForAgencyId = map.get("AgencyID");
			int idxForState = map.get("State");
			int idxForAgencyName = map.get("AgencyName");

			Cell versionCell = dataRow.getCell(idxForVersion);
			Cell agencyIdCell = dataRow.getCell(idxForAgencyId);
			Cell stateCell = dataRow.getCell(idxForState);
			Cell agencyNameCell = dataRow.getCell(idxForAgencyName);

			if (versionCell != null) {
				reportRow.setVersion(String.valueOf(((int) versionCell.getNumericCellValue())));
			}
			if (agencyIdCell != null) {
				reportRow.setAgencyId(agencyIdCell.getStringCellValue());
			}
			if (stateCell != null) {
				reportRow.setState(stateCell.getStringCellValue());
			}
			if (agencyNameCell != null) {
				reportRow.setAgencyName(agencyNameCell.getStringCellValue());
			}

			listOfRdoDataFromReport.add(reportRow);
		}
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

	public Set<String> getStateNames() {
		Set<String> stateList = new LinkedHashSet<>();
		for (TdoReportRow tdoReportRow : listOfRdoDataFromReport) {
			stateList.add(tdoReportRow.getState());
		}
		return stateList;
	}

	public List<String> getAgencyIDs() {
		List<String> agencyList = new ArrayList<>();
		for (TdoReportRow tdoReportRow : listOfRdoDataFromReport) {
			agencyList.add(tdoReportRow.getAgencyId());
		}
		return agencyList;
	}

	public static void main(String[] args) {

		String tdoFilepath = "E:\\client\\files\\TDO2020.xlsx";
		String tdoSheetName = "BC3-YESFORMATV";

		ReadValidateExcel5 readValidateExcel = new ReadValidateExcel5();
		try {
			System.out.println(readValidateExcel.readAndValidateExcelData(tdoFilepath, tdoSheetName));

			// Method to Get state names from excel
			Set<String> stateList = readValidateExcel.getStateNames();

			// Method to Get the list of agency ID's
			List<String> agencyList = readValidateExcel.getAgencyIDs();

			System.out.println("State List =>" + stateList);
			System.out.println("Agency List =>" + agencyList);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public class TdoReportRow {
		private String version;
		private String agencyId;
		private String state;
		private String agencyName;

		public String getVersion() {
			return version;
		}

		public void setVersion(String version) {
			this.version = version;
		}

		public String getAgencyId() {
			return agencyId;
		}

		public void setAgencyId(String agencyId) {
			this.agencyId = agencyId;
		}

		public String getState() {
			return state;
		}

		public void setState(String state) {
			this.state = state;
		}

		public String getAgencyName() {
			return agencyName;
		}

		public void setAgencyName(String agencyName) {
			this.agencyName = agencyName;
		}
	}
}
