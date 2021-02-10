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
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import yahoofinance.Stock;
import yahoofinance.YahooFinance;
import yahoofinance.histquotes.HistoricalQuote;
import yahoofinance.histquotes.Interval;

public class FinanceAPIV1 {

	private Map<String, StockDataRow> stockDataMap = new ConcurrentHashMap<>();

	private void fetchFinanceDataAndWriteToExcel(String financeStockFilepath) throws IOException, FileNotFoundException, InterruptedException {

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
				if (stockCell != null) {
					stockNamesList.add(stockCell.getStringCellValue());
				}
			} else {
				physicalNumberOfRows++;
			}
		}

		String[] stockNamesArr = new String[stockNamesList.size()];
		stockNamesArr = stockNamesList.toArray(stockNamesArr);

		ExecutorService service = Executors.newFixedThreadPool(50);
		for (String stockName : stockNamesList) {
			service.execute(new StockData(stockName));
		}
		System.out.println("Current1 execution completed successfully...");

		Thread.sleep(20000);
		for (String stockName : stockNamesList) {
			service.execute(new CurrentStockData(stockName, "Current2"));
		}
		System.out.println("Current2 execution completed successfully...");

		Thread.sleep(20000);
		for (String stockName : stockNamesList) {
			service.execute(new CurrentStockData(stockName, "Current3"));
		}
		System.out.println("Current3 execution completed successfully...");

		Thread.sleep(20000);
		for (String stockName : stockNamesList) {
			service.execute(new CurrentStockData(stockName, "Current4"));
		}
		System.out.println("Current4 execution completed successfully...");

		Thread.sleep(20000);
		for (String stockName : stockNamesList) {
			service.execute(new CurrentStockData(stockName, "Current5"));
		}
		System.out.println("Current5 execution completed successfully...");

		service.shutdown();
		try {
			service.awaitTermination(Long.MAX_VALUE, TimeUnit.MILLISECONDS);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
				try {
					int idxForStock = map.get("Stock");
					Cell stockCell = dataRow.getCell(idxForStock);
					if (stockCell != null) {
						StockDataRow stockDataRow = stockDataMap.get(stockCell.getStringCellValue().trim());

						int idxForFifthYearPrice = map.get("5 yrs");
						Cell fifthYearPriceCell = dataRow.getCell(idxForFifthYearPrice);
						if (fifthYearPriceCell == null) {
							fifthYearPriceCell = dataRow.createCell(idxForFifthYearPrice);
						}
						fifthYearPriceCell.setCellValue(stockDataRow.getFifthYearPrice());

						int idxForFirstYearPrice = map.get("1 yr");
						Cell firstYearPriceCell = dataRow.getCell(idxForFirstYearPrice);
						if (firstYearPriceCell == null) {
							firstYearPriceCell = dataRow.createCell(idxForFirstYearPrice);
						}
						firstYearPriceCell.setCellValue(stockDataRow.getFirstYearPrice());

						int idxForSixthMonthPrice = map.get("6month");
						Cell sixthMonthPriceCell = dataRow.getCell(idxForSixthMonthPrice);
						if (sixthMonthPriceCell == null) {
							sixthMonthPriceCell = dataRow.createCell(idxForSixthMonthPrice);
						}
						sixthMonthPriceCell.setCellValue(stockDataRow.getSixthMonthPrice());

						int idxForOneMonthPrice = map.get("1 month");
						Cell oneMonthPriceCell = dataRow.getCell(idxForOneMonthPrice);
						if (oneMonthPriceCell == null) {
							oneMonthPriceCell = dataRow.createCell(idxForOneMonthPrice);
						}
						oneMonthPriceCell.setCellValue(stockDataRow.getOneMonthPrice());

						int idxForFiveDaysPrice = map.get("5 days");
						Cell fiveDaysPriceCell = dataRow.getCell(idxForFiveDaysPrice);
						if (fiveDaysPriceCell == null) {
							fiveDaysPriceCell = dataRow.createCell(idxForFiveDaysPrice);
						}
						fiveDaysPriceCell.setCellValue(stockDataRow.getFiveDaysPrice());

						int idxForCurrentOnePrice = map.get("Current1");
						Cell currentOnePriceCell = dataRow.getCell(idxForCurrentOnePrice);
						if (currentOnePriceCell == null) {
							currentOnePriceCell = dataRow.createCell(idxForCurrentOnePrice);
						}
						currentOnePriceCell.setCellValue(stockDataRow.getCurrentOnePrice());

						int idxForCurrentTwoPrice = map.get("Current2");
						Cell currentTwoPriceCell = dataRow.getCell(idxForCurrentTwoPrice);
						if (currentTwoPriceCell == null) {
							currentTwoPriceCell = dataRow.createCell(idxForCurrentTwoPrice);
						}
						currentTwoPriceCell.setCellValue(stockDataRow.getCurrentTwoPrice());

						int idxForCurrentThreePrice = map.get("Current3");
						Cell currentThreePriceCell = dataRow.getCell(idxForCurrentThreePrice);
						if (currentThreePriceCell == null) {
							currentThreePriceCell = dataRow.createCell(idxForCurrentThreePrice);
						}
						currentThreePriceCell.setCellValue(stockDataRow.getCurrentThreePrice());

						int idxForCurrentFourPrice = map.get("Current4");
						Cell currentFourPriceCell = dataRow.getCell(idxForCurrentFourPrice);
						if (currentFourPriceCell == null) {
							currentFourPriceCell = dataRow.createCell(idxForCurrentFourPrice);
						}
						currentFourPriceCell.setCellValue(stockDataRow.getCurrentFourPrice());

						int idxForCurrentFivePrice = map.get("Current5");
						Cell currentFivePriceCell = dataRow.getCell(idxForCurrentFivePrice);
						if (currentFivePriceCell == null) {
							currentFivePriceCell = dataRow.createCell(idxForCurrentFivePrice);
						}
						currentFivePriceCell.setCellValue(stockDataRow.getCurrentFivePrice());

						int idxForRange = map.get("Range");
						Cell rangeCell = dataRow.getCell(idxForRange);
						if (rangeCell == null) {
							rangeCell = dataRow.createCell(idxForRange);
						}
						rangeCell.setCellValue(stockDataRow.getRange());

						int idxForTrend1 = map.get("Trend1");
						Cell trend1Cell = dataRow.getCell(idxForTrend1);
						if (trend1Cell == null) {
							trend1Cell = dataRow.createCell(idxForTrend1);
						}
						trend1Cell.setCellValue(stockDataRow.getTrend1());
						if (stockDataRow.getTrend1().equals("Increased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend1Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend1().equals("Decreased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.RED.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend1Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend1().equals("Equals")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend1Cell.setCellStyle(style);
						}

						int idxForTrend2 = map.get("Trend2");
						Cell trend2Cell = dataRow.getCell(idxForTrend2);
						if (trend2Cell == null) {
							trend2Cell = dataRow.createCell(idxForTrend2);
						}
						trend2Cell.setCellValue(stockDataRow.getTrend2());
						if (stockDataRow.getTrend2().equals("Increased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend2Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend2().equals("Decreased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.RED.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend2Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend2().equals("Equals")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend2Cell.setCellStyle(style);
						}

						int idxForTrend3 = map.get("Trend3");
						Cell trend3Cell = dataRow.getCell(idxForTrend3);
						if (trend3Cell == null) {
							trend3Cell = dataRow.createCell(idxForTrend3);
						}
						trend3Cell.setCellValue(stockDataRow.getTrend3());
						if (stockDataRow.getTrend3().equals("Increased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend3Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend3().equals("Decreased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.RED.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend3Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend3().equals("Equals")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend3Cell.setCellStyle(style);
						}

						int idxForTrend4 = map.get("Trend4");
						Cell trend4Cell = dataRow.getCell(idxForTrend4);
						if (trend4Cell == null) {
							trend4Cell = dataRow.createCell(idxForTrend4);
						}
						trend4Cell.setCellValue(stockDataRow.getTrend4());
						if (stockDataRow.getTrend4().equals("Increased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend4Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend4().equals("Decreased")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.RED.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend4Cell.setCellStyle(style);
						} else if (stockDataRow.getTrend4().equals("Equals")) {
							CellStyle style = workbook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trend4Cell.setCellStyle(style);
						}
					}
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

	public class StockData implements Runnable {
		String stockName;

		public StockData(String stockName) {
			this.stockName = stockName;
		}

		@Override
		public void run() {
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

			Calendar from5Days = Calendar.getInstance();
			from5Days.add(Calendar.DAY_OF_MONTH, -5);

			Calendar to5Days = Calendar.getInstance();

			StockDataRow stockDataRow = new StockDataRow();
			stockDataRow.setStockName(stockName.trim());
			stockDataMap.put(stockDataRow.getStockName().trim(), stockDataRow);

			Stock stock = null;
			try {
				stock = YahooFinance.get(stockName.trim());
			} catch (IOException e1) {
				System.out.println("Exception while fetching stock data from Finance API===>>>" + e1);
				return;
			}

			if (stock != null) {
				try {
					BigDecimal currentPrice = stock.getQuote().getPrice();
					stockDataRow.setCurrentOnePrice(currentPrice.toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API===>>>" + e);
				}

				try {
					List<HistoricalQuote> histQuotes5Days = stock.getHistory(from5Days, to5Days, Interval.DAILY);
					stockDataRow.setFiveDaysPrice(histQuotes5Days.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API for 5 days===>>>" + e);
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
		}
	}

	public class CurrentStockData implements Runnable {
		String stockName;
		String currentRow;

		public CurrentStockData(String stockName, String currentRow) {
			this.stockName = stockName;
			this.currentRow = currentRow;
		}

		@Override
		public void run() {
			StockDataRow stockDataRow = stockDataMap.get(stockName.trim());
			Stock stock = null;
			try {
				stock = YahooFinance.get(stockName.trim());
			} catch (IOException e1) {
				System.out.println("Exception while fetching current stock data from Finance API===>>>" + e1);
				return;
			}

			if (stock != null) {
				try {
					BigDecimal currentPrice = stock.getQuote().getPrice();
					if (currentRow.equals("Current2")) {
						stockDataRow.setCurrentTwoPrice(currentPrice.toString());
					} else if (currentRow.equals("Current3")) {
						stockDataRow.setCurrentThreePrice(currentPrice.toString());
					} else if (currentRow.equals("Current4")) {
						stockDataRow.setCurrentFourPrice(currentPrice.toString());
					} else if (currentRow.equals("Current5")) {
						stockDataRow.setCurrentFivePrice(currentPrice.toString());
						if (stockDataRow.getFifthYearPrice() != null && stockDataRow.getFirstYearPrice() != null) {
							BigDecimal fifthYearPrice = new BigDecimal(stockDataRow.getFifthYearPrice());
							BigDecimal firstYearPrice = new BigDecimal(stockDataRow.getFirstYearPrice());
							int result = firstYearPrice.compareTo(fifthYearPrice);
							if (result >= 1) {
								stockDataRow.setTrend1("Increased");
							} else if (result <= -1) {
								stockDataRow.setTrend1("Decreased");
							} else if (result == 0) {
								stockDataRow.setTrend1("Equals");
							}
						} else {
							stockDataRow.setTrend1("NA");
						}

						if (stockDataRow.getSixthMonthPrice() != null && stockDataRow.getOneMonthPrice() != null) {
							BigDecimal sixthMonthPrice = new BigDecimal(stockDataRow.getSixthMonthPrice());
							BigDecimal oneMonthPrice = new BigDecimal(stockDataRow.getOneMonthPrice());
							int result = oneMonthPrice.compareTo(sixthMonthPrice);
							if (result >= 1) {
								stockDataRow.setTrend2("Increased");
							} else if (result <= -1) {
								stockDataRow.setTrend2("Decreased");
							} else if (result == 0) {
								stockDataRow.setTrend2("Equals");
							}
						} else {
							stockDataRow.setTrend2("NA");
						}

						if (stockDataRow.getOneMonthPrice() != null && stockDataRow.getFiveDaysPrice() != null) {
							BigDecimal oneMonthPrice = new BigDecimal(stockDataRow.getOneMonthPrice());
							BigDecimal fiveDaysPrice = new BigDecimal(stockDataRow.getFiveDaysPrice());
							int result = fiveDaysPrice.compareTo(oneMonthPrice);
							if (result >= 1) {
								stockDataRow.setTrend3("Increased");
							} else if (result <= -1) {
								stockDataRow.setTrend3("Decreased");
							} else if (result == 0) {
								stockDataRow.setTrend3("Equals");
							}
						} else {
							stockDataRow.setTrend3("NA");
						}

						if (stockDataRow.getCurrentOnePrice() != null && stockDataRow.getCurrentFivePrice() != null) {
							BigDecimal currentOnePrice = new BigDecimal(stockDataRow.getCurrentOnePrice());
							BigDecimal currentFivePrice = new BigDecimal(stockDataRow.getCurrentFivePrice());
							int result = currentFivePrice.compareTo(currentOnePrice);
							if (result >= 1) {
								stockDataRow.setTrend4("Increased");
							} else if (result <= -1) {
								stockDataRow.setTrend4("Decreased");
							} else if (result == 0) {
								stockDataRow.setTrend4("Equals");
							}
						} else {
							stockDataRow.setTrend4("NA");
						}

						if (currentPrice != null) {
							BigDecimal big5 = new BigDecimal(5);
							int result5 = currentPrice.compareTo(big5);
							if (result5 <= -1) {
								stockDataRow.setRange("Less than 5");
								return;
							}

							BigDecimal big10 = new BigDecimal(10);
							int result10 = currentPrice.compareTo(big10);
							if (result10 <= -1) {
								stockDataRow.setRange("5 to 10");
								return;
							}

							BigDecimal big20 = new BigDecimal(20);
							int result20 = currentPrice.compareTo(big20);
							if (result20 <= -1) {
								stockDataRow.setRange("10 to 20");
								return;
							}

							BigDecimal big30 = new BigDecimal(30);
							int result30 = currentPrice.compareTo(big30);
							if (result30 <= -1) {
								stockDataRow.setRange("20 to 30");
								return;
							}

							BigDecimal big50 = new BigDecimal(50);
							int result50 = currentPrice.compareTo(big50);
							if (result50 <= -1) {
								stockDataRow.setRange("30 to 50");
								return;
							}

							BigDecimal big100 = new BigDecimal(100);
							int result100 = currentPrice.compareTo(big100);
							if (result100 <= -1 || result100 == 0) {
								stockDataRow.setRange("More than 50");
								return;
							} else if (result100 >= 1) {
								stockDataRow.setRange("More than 100");
								return;
							}
						}
					}
				} catch (Exception e) {
					System.out.println("Exception while fetching current stock data from Finance API===>>>" + e);
				}
			}
		}
	}

	public static void main(String[] args) throws IOException, InterruptedException {

		String financeStockFilepath = "E:\\client\\files\\Data.xlsx";

		FinanceAPIV1 financeAPI = new FinanceAPIV1();

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
		private String fiveDaysPrice;
		private String currentOnePrice;
		private String currentTwoPrice;
		private String currentThreePrice;
		private String currentFourPrice;
		private String currentFivePrice;
		private String range;
		private String trend1;
		private String trend2;
		private String trend3;
		private String trend4;

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

		public String getFiveDaysPrice() {
			return fiveDaysPrice;
		}

		public void setFiveDaysPrice(String fiveDaysPrice) {
			this.fiveDaysPrice = fiveDaysPrice;
		}

		public String getCurrentOnePrice() {
			return currentOnePrice;
		}

		public void setCurrentOnePrice(String currentOnePrice) {
			this.currentOnePrice = currentOnePrice;
		}

		public String getCurrentTwoPrice() {
			return currentTwoPrice;
		}

		public void setCurrentTwoPrice(String currentTwoPrice) {
			this.currentTwoPrice = currentTwoPrice;
		}

		public String getCurrentThreePrice() {
			return currentThreePrice;
		}

		public void setCurrentThreePrice(String currentThreePrice) {
			this.currentThreePrice = currentThreePrice;
		}

		public String getCurrentFourPrice() {
			return currentFourPrice;
		}

		public void setCurrentFourPrice(String currentFourPrice) {
			this.currentFourPrice = currentFourPrice;
		}

		public String getCurrentFivePrice() {
			return currentFivePrice;
		}

		public void setCurrentFivePrice(String currentFivePrice) {
			this.currentFivePrice = currentFivePrice;
		}

		public String getRange() {
			return range;
		}

		public void setRange(String range) {
			this.range = range;
		}

		public String getTrend1() {
			return trend1;
		}

		public void setTrend1(String trend1) {
			this.trend1 = trend1;
		}

		public String getTrend2() {
			return trend2;
		}

		public void setTrend2(String trend2) {
			this.trend2 = trend2;
		}

		public String getTrend3() {
			return trend3;
		}

		public void setTrend3(String trend3) {
			this.trend3 = trend3;
		}

		public String getTrend4() {
			return trend4;
		}

		public void setTrend4(String trend4) {
			this.trend4 = trend4;
		}
	}
}
