package com.client.program;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import yahoofinance.Stock;
import yahoofinance.YahooFinance;
import yahoofinance.histquotes.HistoricalQuote;
import yahoofinance.histquotes.Interval;

public class FinanceAPI {

	private Map<String, StockDataRow> stockDataMap = new HashMap<>();

	private void fetchFinanceDataAndWriteToExcel(String financeStockFilepath) throws IOException, FileNotFoundException {

		if (!new File(financeStockFilepath).exists()) {
			System.out.println("File not exist in directory");
			return;
		}

		Workbook workbook = WorkbookFactory.create(new FileInputStream(financeStockFilepath));
		Sheet sheet = workbook.getSheetAt(0);

		Map<String, Integer> map = new HashMap<String, Integer>();
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		List<String> stockNamesList = new ArrayList<String>();
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				int idxForStock = map.get("Stock");
				Cell stockCell = dataRow.getCell(idxForStock);
				stockNamesList.add(stockCell.getStringCellValue());
			} else {
				physicalNumberOfRows++;
			}
		}

		Calendar from5Year = Calendar.getInstance();
		from5Year.add(Calendar.YEAR, -5);

		Calendar to5Year = Calendar.getInstance();
		to5Year.add(Calendar.YEAR, -5);
		to5Year.add(Calendar.DATE, 1);

		Calendar from1Year = Calendar.getInstance();
		from1Year.add(Calendar.YEAR, -1);

		Calendar to1Year = Calendar.getInstance();
		to1Year.add(Calendar.YEAR, -1);
		to1Year.add(Calendar.MONTH, 1);

		Calendar from6Month = Calendar.getInstance();
		from6Month.add(Calendar.MONTH, -6);

		Calendar to6Month = Calendar.getInstance();
		to6Month.add(Calendar.MONTH, -5);

		Calendar from1Month = Calendar.getInstance();
		from1Month.add(Calendar.MONTH, -1);

		Calendar to1Month = Calendar.getInstance();

		String[] stockNamesArr = new String[stockNamesList.size()];
		stockNamesArr = stockNamesList.toArray(stockNamesArr);

		for (String stockName : stockNamesList) {
			StockDataRow stockDataRow = new StockDataRow();
			stockDataRow.setStockName(stockName.trim());
			stockDataMap.put(stockDataRow.getStockName(), stockDataRow);

			Stock stock = YahooFinance.get(stockName.trim());
			try {
				BigDecimal currentPrice = stock.getQuote().getPrice();
				stockDataRow.setCurrentPrice(currentPrice.toString());
			} catch (Exception e) {
				System.out.println("Exception while fetching stock data from Finance API===>>>" + e);
			}

			try {
				List<HistoricalQuote> histQuotes1Month = stock.getHistory(from1Month, to1Month, Interval.DAILY);
				stockDataRow.setOneMonthPrice(histQuotes1Month.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
			} catch (Exception e) {
				System.out.println("Exception while fetching stock data from Finance API for 1 Month===>>>" + e);
			}

			try {
				List<HistoricalQuote> histQuotes6Month = stock.getHistory(from6Month, to6Month, Interval.DAILY);
				stockDataRow.setSixthMonthPrice(histQuotes6Month.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
			} catch (Exception e) {
				System.out.println("Exception while fetching stock data from Finance API for 6 Months===>>>" + e);
			}

			try {
				List<HistoricalQuote> histQuotes1Year = stock.getHistory(from1Year, to1Year, Interval.DAILY);
				stockDataRow.setFirstYearPrice(histQuotes1Year.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
			} catch (Exception e) {
				System.out.println("Exception while fetching stock data from Finance API for 1 Year===>>>" + e);
			}

			try {
				List<HistoricalQuote> histQuotes5Years = stock.getHistory(from5Year, to5Year, Interval.DAILY);
				stockDataRow.setFifthYearPrice(histQuotes5Years.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
			} catch (Exception e) {
				System.out.println("Exception while fetching stock data from Finance API for 5 Years===>>>" + e);
			}
		}

		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				try {
					int idxForStock = map.get("Stock");
					Cell stockCell = dataRow.getCell(idxForStock);
					StockDataRow stockDataRow = stockDataMap.get(stockCell.getStringCellValue().trim());

					int idxForFifthYearPrice = map.get("5 yrs");
					Cell fifthYearPriceCell = dataRow.getCell(idxForFifthYearPrice);
					fifthYearPriceCell.setCellValue(stockDataRow.getFifthYearPrice());

					int idxForFirstYearPrice = map.get("1 yr");
					Cell firstYearPriceCell = dataRow.getCell(idxForFirstYearPrice);
					firstYearPriceCell.setCellValue(stockDataRow.getFirstYearPrice());

					int idxForSixthMonthPrice = map.get("6month");
					Cell sixthMonthPriceCell = dataRow.getCell(idxForSixthMonthPrice);
					sixthMonthPriceCell.setCellValue(stockDataRow.getSixthMonthPrice());

					int idxForOneMonthPrice = map.get("1 month");
					Cell oneMonthPriceCell = dataRow.getCell(idxForOneMonthPrice);
					oneMonthPriceCell.setCellValue(stockDataRow.getOneMonthPrice());

					int idxForCurrentPrice = map.get("Current");
					Cell currentPriceCell = dataRow.getCell(idxForCurrentPrice);
					currentPriceCell.setCellValue(stockDataRow.getCurrentPrice());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Map===>>>" + e);
				}
			} else {
				physicalNumberOfRows++;
			}
		}

		FileOutputStream fileOut = new FileOutputStream(financeStockFilepath);
		workbook.write(fileOut);
		workbook.close();
		fileOut.close();
	}

	// Include below maven dependency before run the program
	// <dependency>
	// <groupId>com.yahoofinance-api</groupId>
	// <artifactId>YahooFinanceAPI</artifactId>
	// <version>3.15.0</version>
	// </dependency>

	// (If you don't have below jars of POI then please add otherwise ignore)
	// <dependency>
	// <groupId>org.apache.poi</groupId>
	// <artifactId>poi</artifactId>
	// <version>4.1.2</version>
	// </dependency>
	// <dependency>
	// <groupId>org.apache.poi</groupId>
	// <artifactId>poi-ooxml</artifactId>
	// <version>4.1.2</version>
	// </dependency>
	// <dependency>
	// <groupId>org.apache.poi</groupId>
	// <artifactId>poi-ooxml-schemas</artifactId>
	// <version>4.1.2</version>
	// </dependency>
	public static void main(String[] args) throws IOException {

		String financeStockFilepath = "E:\\client\\files\\Shares.xlsx";

		FinanceAPI financeAPI = new FinanceAPI();

		System.out.println("Data fetching and copying into excel started =====>>>");
		financeAPI.fetchFinanceDataAndWriteToExcel(financeStockFilepath);
		System.out.println("Data fetching and copying into excel completed =====>>>");
	}

	public class StockDataRow {
		private String stockName = StringUtils.EMPTY;
		private String fifthYearPrice;
		private String firstYearPrice;
		private String sixthMonthPrice;
		private String oneMonthPrice;
		private String currentPrice;

		public String getStockName() {
			return stockName;
		}

		public void setStockName(String stockName) {
			this.stockName = stockName;
		}

		public String getFifthYearPrice() {
			return fifthYearPrice;
		}

		public void setFifthYearPrice(String fifthYearPrice) {
			this.fifthYearPrice = fifthYearPrice;
		}

		public String getFirstYearPrice() {
			return firstYearPrice;
		}

		public void setFirstYearPrice(String firstYearPrice) {
			this.firstYearPrice = firstYearPrice;
		}

		public String getSixthMonthPrice() {
			return sixthMonthPrice;
		}

		public void setSixthMonthPrice(String sixthMonthPrice) {
			this.sixthMonthPrice = sixthMonthPrice;
		}

		public String getOneMonthPrice() {
			return oneMonthPrice;
		}

		public void setOneMonthPrice(String oneMonthPrice) {
			this.oneMonthPrice = oneMonthPrice;
		}

		public String getCurrentPrice() {
			return currentPrice;
		}

		public void setCurrentPrice(String currentPrice) {
			this.currentPrice = currentPrice;
		}
	}
}
