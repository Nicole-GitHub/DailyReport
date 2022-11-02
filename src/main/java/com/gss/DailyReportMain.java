package com.gss;

import java.io.File;
import java.util.Date;
import java.util.Map;
import java.util.Scanner;

import com.gss.ChkDailyReport.ChkDailyReport;
import com.gss.MonthReport.MaintainList;
import com.gss.MonthReport.MonthReport;
import com.gss.RunDailyReport.RunDailyReport;

public class DailyReportMain {

	public static void main(String[] args) {
		try {

			// 取得jar檔的絕對路徑
//			System.out.println("3:"+ ClassLoader.getSystemResource(""));
//			System.out.println("4:"+ DailyReport.class.getResource(""));//DailyReport.class檔案所在路徑
//			System.out.println("5:"+ DailyReport.class.getResource("/")); // Class包所在路徑,得到的是URL物件,用url.getPath()獲取絕對路徑String
//			System.out.println("6:"+ new File("/").getAbsolutePath());
//			System.out.println("7:"+ System.getProperty("user.dir"));
//			System.out.println("9:"+ System.getProperty("java.class.path"));

			String os = System.getProperty("os.name");
			System.out.println("=== NOW TIME ===> " + new Date());
			System.out.println("===os.name===> " + os);
			
			// 判斷當前執行的啟動方式是IDE還是jar
			boolean isStartupFromJar = new File(DailyReportMain.class.getProtectionDomain().getCodeSource().getLocation().getPath()).isFile();
			System.out.println("isStartupFromJar: " + isStartupFromJar);

			String path = System.getProperty("user.dir") + File.separator; // Jar
			if(!isStartupFromJar) // IDE
				path = os.contains("Mac") ? "/Users/nicole/Dropbox/DailyReport/" // Mac
						: "C:/Users/Nicole/Dropbox/DailyReport/"; // win
			
			System.out.println("path: " + path);
			Map<String, String> mapProp = Property.getProperties(path);

			// 執行類別
			String runType = mapProp.get("runType");
			System.out.println("執行類別: " + runType);

			// 月報放置路徑與檔名
			String monthReportPath = "", maintainListExcel = "", monthReportExcel = "";
			int lastMonthTotalFinishNum = 0, lastMonthTotalCodeNum = 0, lastMonth = 0;
			
			if (runType.equals("month") || runType.equals("maintain")) { // 驗証月報內容

				if (!isStartupFromJar) {// IDE
					lastMonthTotalFinishNum = 276;
					lastMonthTotalCodeNum = 200;
					lastMonth = 202210;
				}
				
				/**
				 * 透過windows的cmd執行時需將System.in格式轉為big5才不會讓中文變亂碼
				 * 即使在cmd下chcp 65001轉成utf-8也沒用
				 * 但在eclipse執行時不能轉為big5
				 */
				try (Scanner s =  isStartupFromJar ? new Scanner(System.in, "big5") : new Scanner(System.in)) {
					System.out.println("請輸入 月報月份(yyyymm): ");
					lastMonth = lastMonth == 0 ? s.nextInt() : lastMonth;
					if (runType.equals("month")) {
						System.out.println("請輸入 上月累計已完成件數: ");
						lastMonthTotalFinishNum = lastMonthTotalFinishNum == 0 ? s.nextInt() : lastMonthTotalFinishNum;
						System.out.println("請輸入 上月累計程式支數: ");
						lastMonthTotalCodeNum = lastMonthTotalCodeNum == 0 ? s.nextInt() : lastMonthTotalCodeNum;
					}
				}
				
				// 月報放置路徑與檔名
				monthReportPath = mapProp.get("MonthReportPath") + lastMonth + "/";
				maintainListExcel = path + monthReportPath + mapProp.get("MaintainListExcel") + lastMonth + ".xlsx";
				monthReportExcel = path + monthReportPath + mapProp.get("MonthReportExcel") + (lastMonth-191100) + ".xlsx";
				System.out.println("monthReportPath: " + monthReportPath);
				System.out.println("維護問題紀錄單Excel: " + maintainListExcel);
				System.out.println("月報Excel: " + monthReportExcel);
				
			}
			
			if (runType.equals("check")) {
				ChkDailyReport.chkDailyReport(path); // 檢查日誌
				
			} else if (runType.equals("maintain")) {
				MaintainList.maintainList(maintainListExcel); // 整理維護問題紀錄單Excel
				
			} else if (runType.equals("month")) { // 驗証月報內容
				MonthReport.monthReport(monthReportExcel, lastMonthTotalFinishNum, lastMonthTotalCodeNum); // 驗証月報
				
			} else if (runType.equals("run")){
				/**
				 * 整理日誌
				 * 當錯誤原因為找不到DailyReportExcel檔及ChromeDriver版本錯誤時
				 * 則不再重跑
				 */
				boolean done = false;
				do {
					try {
						RunDailyReport.runDailyReport(path);
						done = true;
					} catch (Exception e) {
						System.out.println(new Date() + " ===> " + e.getMessage());
						if ("getDailyReportExcel Error".equals(e.getMessage())
								|| e.getMessage().contains("This version of ChromeDriver only supports Chrome version"))
							done = true;
					}
				} while (!done);
				
			} else {
				System.out.println("runType Error");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
