package com.client.program;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import yahoofinance.Stock;
import yahoofinance.YahooFinance;
import yahoofinance.histquotes.HistoricalQuote;
import yahoofinance.histquotes.Interval;

public class FinanceAPIV2 {

	private Map<String, HashMap<String, String>> stockDataMap = new ConcurrentHashMap<>();

	private void fetchFinanceDataAndWriteToExcel(String financeStockFilepath, int totalCurrent, int delayInSeconds, String targetFilePath) throws IOException, FileNotFoundException, InterruptedException {

		if (!new File(financeStockFilepath).exists()) {
			System.out.println("File not exist in directory");
			return;
		}

		Workbook workbook = WorkbookFactory.create(new FileInputStream(financeStockFilepath));
		XSSFWorkbook myWorkBook = new XSSFWorkbook();
		Sheet sheet = workbook.getSheetAt(0);
		XSSFSheet mySheet = myWorkBook.createSheet(sheet.getSheetName());

		Map<String, Integer> map = new HashMap<String, Integer>();
		Row row = sheet.getRow(0);

		int minColIx = row.getFirstCellNum();
		int maxColIx = row.getLastCellNum();

		int idxForStock = 0;
		for (int colIx = minColIx; colIx < maxColIx; colIx++) {
			Cell cell = row.getCell(colIx);
			if (cell.getStringCellValue().trim().equals("Stock")) {
				idxForStock = cell.getColumnIndex();
				break;
			}
		}

		List<String> list = new ArrayList<>();
		list.add("Stock");
		list.add("5 yrs");
		list.add("1 yr");
		list.add("6 month");
		list.add("1 month");
		list.add("5 days");
		for (int i = 1; i <= totalCurrent; i++) {
			list.add("Current" + i);
		}
		list.add("Range");
		list.add("Trend1");
		list.add("Trend2");
		list.add("Trend3");
		list.add("Trend4");
		list.add("per1");
		list.add("per2");

		XSSFRow myRow = mySheet.createRow(0);
		for (int i = 0; i < list.size(); i++) {
			XSSFCell myCell = myRow.createCell(i);
			myCell.setCellType(CellType.STRING);
			myCell.setCellValue(list.get(i));

			CellStyle style = myWorkBook.createCellStyle();
			style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			myCell.setCellStyle(style);
		}

		int myMinColIx = myRow.getFirstCellNum();
		int myMaxColIx = myRow.getLastCellNum();

		for (int colIx = myMinColIx; colIx < myMaxColIx; colIx++) {
			Cell cell = myRow.getCell(colIx);
			map.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
		}

		List<String> stockNamesList = new ArrayList<String>();
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		for (int x = 1; x < physicalNumberOfRows; x++) {
			Row dataRow = sheet.getRow(x);
			if (dataRow != null) {
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

		for (int i = 2; i <= totalCurrent; i++) {
			Thread.sleep(delayInSeconds);
			for (String stockName : stockNamesList) {
				service.execute(new CurrentStockData(stockName, "current" + i, totalCurrent));
			}
			System.out.println("Current" + i + " execution completed successfully...");
		}

		service.shutdown();
		try {
			service.awaitTermination(Long.MAX_VALUE, TimeUnit.MILLISECONDS);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}

		for (int x = 0; x < stockNamesList.size(); x++) {
			int newRow = x + 1;
			Row dataRow = mySheet.createRow(newRow);
			try {
				HashMap<String, String> stockDataRow = stockDataMap.get(stockNamesList.get(x).trim());
				int myIdxForStock = map.get("Stock");
				Cell stockCell = dataRow.createCell(myIdxForStock);
				stockCell.setCellValue(stockNamesList.get(x));

				int idxForFifthYearPrice = map.get("5 yrs");
				Cell fifthYearPriceCell = dataRow.createCell(idxForFifthYearPrice);
				if (stockDataRow.get("fifthYearPrice") != null) {
					fifthYearPriceCell.setCellValue(stockDataRow.get("fifthYearPrice").toString());
				}

				int idxForFirstYearPrice = map.get("1 yr");
				Cell firstYearPriceCell = dataRow.createCell(idxForFirstYearPrice);
				if (stockDataRow.get("firstYearPrice") != null) {
					firstYearPriceCell.setCellValue(stockDataRow.get("firstYearPrice").toString());
				}

				int idxForSixthMonthPrice = map.get("6 month");
				Cell sixthMonthPriceCell = dataRow.createCell(idxForSixthMonthPrice);
				if (stockDataRow.get("sixthMonthPrice") != null) {
					sixthMonthPriceCell.setCellValue(stockDataRow.get("sixthMonthPrice").toString());
				}

				int idxForOneMonthPrice = map.get("1 month");
				Cell oneMonthPriceCell = dataRow.createCell(idxForOneMonthPrice);
				if (stockDataRow.get("oneMonthPrice") != null) {
					oneMonthPriceCell.setCellValue(stockDataRow.get("oneMonthPrice").toString());
				}

				int idxForFiveDaysPrice = map.get("5 days");
				Cell fiveDaysPriceCell = dataRow.createCell(idxForFiveDaysPrice);
				if (stockDataRow.get("fiveDaysPrice") != null) {
					fiveDaysPriceCell.setCellValue(stockDataRow.get("fiveDaysPrice").toString());
				}

				for (int i = 1; i <= totalCurrent; i++) {
					int idxForCurrentOnePrice = map.get("Current" + i);
					Cell currentOnePriceCell = dataRow.createCell(idxForCurrentOnePrice);
					if (stockDataRow.get("current" + i + "Price") != null) {
						currentOnePriceCell.setCellValue(stockDataRow.get("current" + i + "Price").toString());
					}
				}

				int idxForRange = map.get("Range");
				Cell rangeCell = dataRow.createCell(idxForRange);
				if (stockDataRow.get("range") != null) {
					rangeCell.setCellValue(stockDataRow.get("range").toString());
				}

				for (int i = 1; i <= 4; i++) {
					int idxForTrend = map.get("Trend" + i);
					Cell trendCell = dataRow.createCell(idxForTrend);
					if (stockDataRow.get("trend" + i) != null) {
						trendCell.setCellValue(stockDataRow.get("trend" + i).toString());
						if (stockDataRow.get("trend" + i).toString().equals("Increased")) {
							CellStyle style = myWorkBook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trendCell.setCellStyle(style);
						} else if (stockDataRow.get("trend" + i).toString().equals("Decreased")) {
							CellStyle style = myWorkBook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.RED.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trendCell.setCellStyle(style);
						} else if (stockDataRow.get("trend" + i).toString().equals("Equals")) {
							CellStyle style = myWorkBook.createCellStyle();
							style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
							style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
							trendCell.setCellStyle(style);
						}
					}
				}

				int idxForPer1 = map.get("per1");
				Cell per1Cell = dataRow.createCell(idxForPer1);
				if (stockDataRow.get("per1") != null) {
					per1Cell.setCellValue(stockDataRow.get("per1").toString());
				}

				int idxForPer2 = map.get("per2");
				Cell per2Cell = dataRow.createCell(idxForPer2);
				if (stockDataRow.get("per2") != null) {
					per2Cell.setCellValue(stockDataRow.get("per2").toString());
				}
			} catch (Exception e) {
				System.out.println("Exception while fetching stock data from Map===>>>" + e);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(new File(targetFilePath));
		myWorkBook.write(fileOut);
		workbook.close();
		myWorkBook.close();
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

			HashMap<String, String> stockDataRow = new HashMap<String, String>();
			stockDataRow.put("stockName", stockName.trim());
			stockDataMap.put(stockDataRow.get("stockName").toString().trim(), stockDataRow);

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
					stockDataRow.put("current1Price", currentPrice.toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API for Current Price===>>>" + e);
				}

				try {
					List<HistoricalQuote> histQuotes5Days = stock.getHistory(from5Days, to5Days, Interval.DAILY);
					stockDataRow.put("fiveDaysPrice", histQuotes5Days.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API for 5 days===>>>" + e);
				}

				try {
					List<HistoricalQuote> histQuotes1Month = stock.getHistory(from1Month, to1Month, Interval.DAILY);
					stockDataRow.put("oneMonthPrice", histQuotes1Month.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API for 1 Month===>>>" + e);
				}

				try {
					List<HistoricalQuote> histQuotes6Month = stock.getHistory(from6Month, to6Month, Interval.DAILY);
					stockDataRow.put("sixthMonthPrice", histQuotes6Month.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API for 6 Months===>>>" + e);
				}

				try {
					List<HistoricalQuote> histQuotes1Year = stock.getHistory(from1Year, to1Year, Interval.DAILY);
					stockDataRow.put("firstYearPrice", histQuotes1Year.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API for 1 Year===>>>" + e);
				}

				try {
					List<HistoricalQuote> histQuotes5Years = stock.getHistory(from5Year, to5Year, Interval.DAILY);
					stockDataRow.put("fifthYearPrice", histQuotes5Years.get(0).getClose().setScale(2, BigDecimal.ROUND_HALF_UP).toString());
				} catch (Exception e) {
					System.out.println("Exception while fetching stock data from Finance API for 5 Years===>>>" + e);
				}
			}
		}
	}

	public class CurrentStockData implements Runnable {
		String stockName;
		String currentRow;
		int totalCurrent;

		public CurrentStockData(String stockName, String currentRow, int totalCurrent) {
			this.stockName = stockName;
			this.currentRow = currentRow;
			this.totalCurrent = totalCurrent;
		}

		@Override
		public void run() {
			HashMap<String, String> stockDataRow = stockDataMap.get(stockName.trim());
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
					if (currentPrice != null) {
						stockDataRow.put(currentRow + "Price", currentPrice.toString());
					}
					if (currentRow.equals("current" + totalCurrent)) {
						if (currentPrice != null) {
							stockDataRow.put(currentRow + "Price", currentPrice.toString());
						}
						if (stockDataRow.get("fifthYearPrice") != null && stockDataRow.get("firstYearPrice") != null) {
							BigDecimal fifthYearPrice = new BigDecimal(stockDataRow.get("fifthYearPrice"));
							BigDecimal firstYearPrice = new BigDecimal(stockDataRow.get("firstYearPrice"));
							int result = firstYearPrice.compareTo(fifthYearPrice);
							if (result >= 1) {
								stockDataRow.put("trend1", "Increased");
							} else if (result <= -1) {
								stockDataRow.put("trend1", "Decreased");
							} else if (result == 0) {
								stockDataRow.put("trend1", "Equals");
							}
						} else {
							stockDataRow.put("trend1", "NA");
						}

						if (stockDataRow.get("sixthMonthPrice") != null && stockDataRow.get("oneMonthPrice") != null) {
							BigDecimal sixthMonthPrice = new BigDecimal(stockDataRow.get("sixthMonthPrice"));
							BigDecimal oneMonthPrice = new BigDecimal(stockDataRow.get("oneMonthPrice"));
							int result = oneMonthPrice.compareTo(sixthMonthPrice);
							if (result >= 1) {
								stockDataRow.put("trend2", "Increased");
							} else if (result <= -1) {
								stockDataRow.put("trend2", "Decreased");
							} else if (result == 0) {
								stockDataRow.put("trend2", "Equals");
							}
						} else {
							stockDataRow.put("trend2", "NA");
						}

						if (stockDataRow.get("oneMonthPrice") != null && stockDataRow.get("fiveDaysPrice") != null) {
							BigDecimal oneMonthPrice = new BigDecimal(stockDataRow.get("oneMonthPrice"));
							BigDecimal fiveDaysPrice = new BigDecimal(stockDataRow.get("fiveDaysPrice"));
							int result = fiveDaysPrice.compareTo(oneMonthPrice);
							if (result >= 1) {
								stockDataRow.put("trend3", "Increased");
							} else if (result <= -1) {
								stockDataRow.put("trend3", "Decreased");
							} else if (result == 0) {
								stockDataRow.put("trend3", "Equals");
							}
						} else {
							stockDataRow.put("trend3", "NA");
						}

						if (stockDataRow.get("current1Price") != null && stockDataRow.get("current" + totalCurrent + "Price") != null) {
							BigDecimal currentOnePrice = new BigDecimal(stockDataRow.get("current1Price"));
							BigDecimal currentLastPrice = new BigDecimal(stockDataRow.get("current" + totalCurrent + "Price"));
							int result = currentLastPrice.compareTo(currentOnePrice);
							if (result >= 1) {
								stockDataRow.put("trend4", "Increased");
							} else if (result <= -1) {
								stockDataRow.put("trend4", "Decreased");
							} else if (result == 0) {
								stockDataRow.put("trend4", "Equals");
							}

							DecimalFormat decimalFormat = new DecimalFormat("#.##");
							float lastPrice = Float.valueOf(decimalFormat.format(Float.parseFloat(stockDataRow.get("current" + totalCurrent + "Price"))));
							float firstPrice = Float.valueOf(decimalFormat.format(Float.parseFloat(stockDataRow.get("current1Price"))));
							if (stockDataRow.get("fiveDaysPrice") != null) {
								float fiveDaysPrice = Float.valueOf(decimalFormat.format(Float.parseFloat(stockDataRow.get("fiveDaysPrice"))));
								float per1 = Float.valueOf(decimalFormat.format(lastPrice * 100 / fiveDaysPrice));
								stockDataRow.put("per1", String.valueOf(per1));
							}

							float per2 = Float.valueOf(decimalFormat.format(lastPrice * 100 / firstPrice));
							stockDataRow.put("per2", String.valueOf(per2));
						} else {
							stockDataRow.put("trend4", "NA");
						}

						stockDataRow.put("range", "NA");
						if (currentPrice != null) {
							BigDecimal big2 = new BigDecimal(2);
							int result2 = currentPrice.compareTo(big2);
							if (result2 <= -1) {
								stockDataRow.put("range", "Less than 2");
								return;
							}

							BigDecimal big5 = new BigDecimal(5);
							int result5 = currentPrice.compareTo(big5);
							if (result5 <= -1) {
								stockDataRow.put("range", "3 to 5");
								return;
							}

							BigDecimal big10 = new BigDecimal(10);
							int result10 = currentPrice.compareTo(big10);
							if (result10 <= -1) {
								stockDataRow.put("range", "5 to 10");
								return;
							}

							BigDecimal big20 = new BigDecimal(20);
							int result20 = currentPrice.compareTo(big20);
							if (result20 <= -1) {
								stockDataRow.put("range", "10 to 20");
								return;
							}

							BigDecimal big30 = new BigDecimal(30);
							int result30 = currentPrice.compareTo(big30);
							if (result30 <= -1) {
								stockDataRow.put("range", "20 to 30");
								return;
							}

							BigDecimal big50 = new BigDecimal(50);
							int result50 = currentPrice.compareTo(big50);
							if (result50 <= -1) {
								stockDataRow.put("range", "30 to 50");
								return;
							}

							BigDecimal big100 = new BigDecimal(100);
							int result100 = currentPrice.compareTo(big100);
							if (result100 <= -1 || result100 == 0) {
								stockDataRow.put("range", "More than 50");
								return;
							} else if (result100 >= 1) {
								stockDataRow.put("range", "More than 100");
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
		String targetFilePath;
		int totalCurrent = 10;
		int delayInSeconds = 20;

		FileSystem system = FileSystems.getDefault();
		Path original = system.getPath(financeStockFilepath);

		String origFileName = original.getFileName().toString();
		String fileName = origFileName.substring(0, origFileName.lastIndexOf("."));
		String fileExt = origFileName.substring(origFileName.lastIndexOf("."), origFileName.length());
		Path targetFile = system.getPath(original.getParent() + "\\" + fileName + "_"+ new Date().toString().replace(" ", "_").replace(":", "_") + fileExt);
		targetFilePath = targetFile.toString();

		FinanceAPIV2 financeAPI = new FinanceAPIV2();

		System.out.println("Data fetching and copying into excel started =====>>>");
		financeAPI.fetchFinanceDataAndWriteToExcel(financeStockFilepath, totalCurrent, delayInSeconds * 1000, targetFilePath);
		System.out.println("Data fetching and copying into excel completed =====>>>");
	}
}
