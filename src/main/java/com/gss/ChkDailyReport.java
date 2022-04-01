package com.gss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ChkDailyReport {

	protected static void chkDailyReport(String path) {
		XSSFWorkbook xssfWorkbook = null;
		OutputStream output = null;
		String minRunDate = "", maxRunDate = "", minRunDay = "";
		int minRunMonth = 0, maxRunMonth = 0, minusDays = 0, dateCell = 0;
		Cell targetCell;

//		path = "/Users/nicole/22/java/eclipse-workspace/DailyReport/src/main/java/com/gss/DailyReport/"; // Debug
		System.out.println("path: " + path);

		Map<String, String> mapProp = Property.getProperties(path);
//		path = "/Users/nicole/Downloads/"; // Debug
		// 日報放置路徑與檔名
		String ChkDailyReportExcel = mapProp.get("ChkDailyReportExcel"); // Jar
		String DailyReportExcel = path + ChkDailyReportExcel; // Jar
		System.out.println("日報Excel: " + DailyReportExcel);
		String excelDate = ChkDailyReportExcel.substring(ChkDailyReportExcel.lastIndexOf("_") + 1,
				ChkDailyReportExcel.lastIndexOf("."));
		String excelYear = excelDate.substring(0, 4);
		String excelMonth = excelDate.substring(4, 6);
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
		Calendar cal = Calendar.getInstance();
		/**
		 * 設定字串轉日期 cal.set(年，月，日) 月的初始值為0 日若為0則為前一個月的最後一天
		 */
		cal.set(Integer.parseInt(excelYear), Integer.parseInt(excelMonth), 0);
		System.out.println("cal1: " + sdf.format(cal.getTime()));
		String lastday = sdf.format(cal.getTime()).substring(6);

		try {
			File f = new File(DailyReportExcel);
			InputStream inputStream = new FileInputStream(f);
			// XSSF (.xlsx)
			xssfWorkbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet1 = xssfWorkbook.getSheetAt(0);

			// 取得對應Job的Row位置(橫列)
			for (Row row : sheet1) {
				if (row.getRowNum() > 1 && row.getCell(0) != null && row.getRowNum() < sheet1.getLastRowNum()) {

					System.out.println("getRowNum:" + row.getRowNum());
					// 最小執行日期 (yyyymmdd)
					minRunDate = row.getCell(4).getStringCellValue();
					// 最大執行日期 (yyyymmdd)
					maxRunDate = row.getCell(5).getStringCellValue();
					// 最小執行日期 (yyyymm)
					minRunMonth = Integer.parseInt(minRunDate.substring(1, 7));
					// 最大執行日期 (yyyymm)
					maxRunMonth = Integer.parseInt(maxRunDate.substring(1, 7));
					if (minRunMonth == maxRunMonth) {
						// 最大執行日期 - 最小執行日期
						minusDays = Integer.parseInt(maxRunDate.substring(1))
								- Integer.parseInt(minRunDate.substring(1));
						// 最小執行日期 (dd)
						minRunDay = minRunDate.substring(7);

						// 最大執行日期(yyyymm) = Excel年月
					} else if (String.valueOf(maxRunMonth).equals(excelYear + excelMonth)) {
						minRunDay = "1";
						minusDays = Integer.parseInt(maxRunDate.substring(7)) - Integer.parseInt(minRunDay);

						// 最大執行日期為隔月
					} else {
						minRunDay = minRunDate.substring(7);
						// Excel年月的最後一日 - 最小執行日期 (dd)
						minusDays = Integer.parseInt(lastday) - Integer.parseInt(minRunDate.substring(7));
					}

					// 取得對應日期的Cell位置(縱列)
					dateCell = Tools.getDateCell(sheet1, minRunDay);
					for (int i = 0; i <= minusDays; i++) {
						row.getCell(dateCell).setCellValue("V");
						dateCell += 2;
						System.out.println("getRowNum:" + row.getRowNum() + ", getCell:" + dateCell);
					}
				} else if (row != null && row.getRowNum() == sheet1.getLastRowNum()) {
					// 取得對應日期的Cell位置(縱列)
					dateCell = Tools.getDateCell(sheet1, "1");
					for (int i = 0; i < Integer.parseInt(lastday); i++) {
						targetCell = row.getCell(dateCell);
						// 需再執行一次有公式的欄位才會更新欄位值
						if (targetCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
							targetCell.setCellFormula(targetCell.getCellFormula());
							dateCell += 2;
						}
					}
				}
			}

			output = new FileOutputStream(f);
			xssfWorkbook.write(output);
		} catch (Exception ex) {
			System.out.println("Error:");
			ex.printStackTrace();
		} finally {
			try {
				if(xssfWorkbook != null)
					xssfWorkbook.close();
				if(output != null)
					output.close();
			} catch (IOException e) {
				System.out.println("Error:" + e.getMessage());
			}
		}
		System.out.println("Done!");
	}
}
