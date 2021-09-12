package com.gss;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class RunDailyReport {
	static final String yy1 = "20"; // 西元年前兩碼
	static final String mailStartText = " - 您好, 〔("; // job 寄的 mail 開頭
	static final SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
	static Integer dateCell = 0, dataRow = 0, chkDate = 0, mailItemInit = 0; // 從第0封mail開始取;
	static String JobMonth = "", JobDate = "", excelMonth = "", DailyReportExcel = "", account = "", pwd = "";
	static boolean isPrint;
	static Row targetRow;
	static Cell targetCell, previouCell, targetChkCell, previouChkCell;
	static ArrayList<Map<String, String>> listF, list;
	static String[] inboxName;

	/**
	 * 整理日誌
	 * 
	 * @throws IOException
	 * @throws ParseException
	 */
	protected static void runDailyReport(String path) {

		Map<String, String> mapProp = Property.getProperties(path);
		// 要從哪天的mail開始取
		chkDate = Integer.parseInt(mapProp.get("chkDate"));
		System.out.println("要從 " + chkDate + " 的mail開始取 ");

		// 日報放置路徑與檔名
		DailyReportExcel = path + mapProp.get("DailyReportExcel"); // Jar
		System.out.println("日報Excel: " + DailyReportExcel);

		// 收件匣名稱
		inboxName = mapProp.get("inboxName").split(",");
		System.out.println("收件匣名稱: " + inboxName);

		// 帳號
		account = mapProp.get("account");
		System.out.println("Mail帳號: " + account);

		// 密碼
		pwd = mapProp.get("pwd");
		System.out.println("Mail密碼: " + pwd);

		Workbook workbook = null;
		OutputStream output = null;

		try {
			// 整理 MAIL內容
			mailContent(path);

			File f = new File(DailyReportExcel);
			workbook = Tools.getWorkbook(DailyReportExcel, f);
			Sheet sheet1 = workbook.getSheetAt(0);

			// 日誌的月份 年月(六碼)
			excelMonth = sheet1.getRow(0).getCell(0).getStringCellValue();
			excelMonth = excelMonth.substring(1, 7);

			// 寫入 "JobList" 頁籤的狀態，並整理出失敗的Job
			writeSheet1(sheet1);

			// 將失敗的job列進 "待辦JOB" 頁籤
			writeSheet3(workbook);

			System.out.println("Done!");

			output = new FileOutputStream(f);
			workbook.write(output);
		} catch (Exception ex) {
			System.out.println("runDailyReport catch Error:");
			ex.printStackTrace();
		} finally {
			try {
				if (workbook != null)
					workbook.close();
				if (output != null)
					output.close();
			} catch (IOException ex) {
				System.out.println("runDailyReport finally Error:");
				ex.printStackTrace();
			}
		}
	}

	/**
	 * 整理 MAIL內容
	 * 
	 * @return 整理後的list
	 *         List<Map<String, String>>
	 *         map.put("jobRSDate", jobRSDate);
	 *         map.put("jobRSTime", jobRSTime);
	 *         map.put("jobRSDateTime", jobRSDate + jobRSTime);
	 *         map.put("jobRSOriDateTime", jobRSOriDate + jobRSTime);
	 *         map.put("jobPeriod", jobPeriod); map.put("jobSeq", jobSeq);
	 *         map.put("jobEName", jobEName); map.put("jobName", jobName);
	 *         map.put("jobRunRS", jobRunRS);
	 * @throws ParseException
	 */
	private static void mailContent(String path) throws ParseException {
		boolean isPm = false;
		String jobRSText = "", jobRSDate = "", jobRSOriDate = "", jobRSTime = "",
				jobPeriod = "", jobSeq = "", jobEName = "", jobName = "",
				jobRunRS = "", yy2 = "", mm = "", dd = "", hh = "", mi = "", time = "";
		Map<String, String> map;
		list = new ArrayList<Map<String, String>>();
		List<Map<String, String>> listMail;

		// 取 昨、今、明 三天的日期
		Calendar cal = Calendar.getInstance();
		String today = sdf.format(cal.getTime());
		cal.add(Calendar.DATE, -1);
		String yesterday = sdf.format(cal.getTime());
		cal.add(Calendar.DATE, +2);
		String tomorrow = sdf.format(cal.getTime());
		String[] jobMailTitleArr, jobRSDateArr;
		int arrLen = 0;

		cal.setTime(sdf.parse(String.valueOf(chkDate)));
		cal.add(Calendar.DATE, -2);
		SimpleDateFormat mailSDF = new SimpleDateFormat("yy/M/d");
		String mailScrolltoDate = mailSDF.format(cal.getTime());

		// maillist
		listMail = Selenium_Crawler.getMailContent(path, inboxName, account, pwd, mailScrolltoDate);

		for (Map<String, String> mailMap : listMail) {

			System.out.println(mailMap.get("title"));
			System.out.println(mailMap.get("body"));
			// mail內容
			jobRSText = mailMap.get("body");

			// 是否為 job 的 mail
			if (jobRSText.indexOf(mailStartText) == 0) {

				/**
				 * jobMailTitleArr 最後一個逗號有時會多空隔有時不會，因此先將空格移除再做split
				 * 0: 已讀
				 * 1: NHIA
				 * 2: 有附件
				 * 3: 執行成功=>(386062)彙總－疾病就醫利用彙整檔(補歷史資料)(DWF)_2007/01(附檔)
				 * 4: 收件匣/NHIA/2JOB成功
				 * 5: 18KB
				 * 6: "今天" or "21/1/16六"
				 * 7: 上午10:22
				 */
				jobMailTitleArr = mailMap.get("title").replace(" ", "").split(",");
				// 有時job的中文名稱會有逗號，會影響到陣列的總數
				arrLen = jobMailTitleArr.length;

				/**
				 * mail 收到的實際時間
				 * 格式 年月日(各兩碼)
				 * 切割時間點為上午9點
				 * (9:00前屬當天，9:00後屬隔天)
				 */
				time = jobMailTitleArr[arrLen - 1].trim();
				isPm = time.substring(0, 2).equals("下午");
				time = time.substring(2);
				hh = time.substring(0, time.indexOf(":"));
				hh = hh.length() < 2 ? "0" + hh : hh; // 不滿兩位前面補0
				hh = "12".equals(hh) ? "00" : hh;
				mi = time.substring(time.indexOf(":") + 1);
				// mail收到的時間 (上午:1,下午:2)
				jobRSTime = (isPm ? "2" : "1") + hh + mi;
				// job 執行區間
//				jobPeriod = jobMailTitleArr[4].startsWith("收件匣") ? jobMailTitleArr[3] : jobMailTitleArr[4];
				jobPeriod = jobMailTitleArr[arrLen - 5]; // 因job中文名稱內可能會有多個逗號，故抓倒數第五位陣列值
				jobPeriod = jobPeriod.substring(jobPeriod.lastIndexOf("_") + 1, jobPeriod.lastIndexOf("("));
				// job 所屬日期 (日誌用)
				jobRSDate = jobMailTitleArr[arrLen - 2].trim();
				// job 原日期 (後面刪除相同job時使用)
				jobRSOriDate = jobRSDate.equals("昨天") ? yesterday : jobRSDate.equals("今天") ? today : "";
				if (jobRSDate.lastIndexOf("/") > 0) {
					jobRSDateArr = jobRSDate.split("/");
					yy2 = jobRSDateArr[0];
					mm = jobRSDateArr[1];
					dd = jobRSDateArr[2].substring(0, jobRSDateArr[2].length() - 1);
					jobRSOriDate = yy1 + yy2 + Tools.getLen2(mm) + Tools.getLen2(dd);
					/**
					 * "上午 9:00" 前屬當天
					 * "上午 9:00" 後屬隔天
					 */
					if (isPm || (!isPm && Integer.valueOf(hh) > 8)) {
						dd = String.valueOf((Integer.valueOf(dd) + 1));
					}
					jobRSDate = yy1 + yy2 + Tools.getLen2(mm) + Tools.getLen2(dd);

					// 時間點為 "昨天, 上午 9:00" 前
				} else if (jobRSDate.equals("昨天") && !isPm && Integer.valueOf(hh) <= 8) {
					jobRSDate = yesterday;

					/**
					 * 時間點為 "今天, 上午 9:00" 前
					 * 或 "昨天, 下午"
					 * 或 "昨天, 上午 9:00" 後
					 */
				} else if ((jobRSDate.equals("今天") && !isPm && Integer.valueOf(hh) <= 8)
						|| (jobRSDate.equals("昨天") && (isPm || (!isPm && Integer.valueOf(hh) > 8)))) {
					jobRSDate = today;
				} else {
					jobRSDate = tomorrow;
				}

				// 是否為這次要整理的job日期範圍
				isPrint = Integer.parseInt(jobRSDate) >= chkDate;
				// Job Seq
				jobSeq = jobRSText.substring(jobRSText.indexOf("(") + 1, jobRSText.indexOf(")"));
				// job 英文名
				jobEName = jobRSText.substring(jobRSText.indexOf(")") + 1, jobRSText.indexOf("-", 2));
				// job 中文名 (不是太重要，所以尾部還有些符號與類別沒處理)
				jobName = jobRSText.substring(jobRSText.indexOf("-", 2) + 1, jobRSText.lastIndexOf("於 "));
				jobName = jobName.substring(0, jobName.lastIndexOf("〕〔"));
				// job 執行結果
				jobRunRS = jobRSText.substring(jobRSText.lastIndexOf("之排程工作執行") + 7, jobRSText.lastIndexOf("。"));
				jobRunRS = jobRunRS.equals("成功") ? "S" : jobRunRS.equals("失敗") ? "F" : "Z";

				if (isPrint) {
					map = new HashMap<String, String>();
					map.put("jobRSDate", jobRSDate);
					map.put("jobRSTime", jobRSTime);
					map.put("jobRSDateTime", jobRSDate + jobRSTime);
					map.put("jobRSOriDateTime", jobRSOriDate + jobRSTime);
					map.put("jobPeriod", jobPeriod);
					map.put("jobSeq", jobSeq);
					map.put("jobEName", jobEName);
					map.put("jobName", jobName);
					map.put("jobRunRS", jobRunRS);
					list.add(map);
//					System.out.println("======================== Start ========================");
//					System.out.println("jobRSDate=>"+jobRSDate);
//					System.out.println("jobRSTime=>"+ jobRSTime);
//					System.out.println("jobRSDateTime=>"+ jobRSDate + jobRSTime);
//					System.out.println("jobRSOriDateTime=>"+ jobRSOriDate + jobRSTime);
//					System.out.println("jobPeriod=>"+ jobPeriod);
//					System.out.println("jobSeq=>"+ jobSeq);
//					System.out.println("jobEName=>"+ jobEName);
//					System.out.println("jobName=>"+ jobName);
//					System.out.println("jobRunRS=>"+ jobRunRS);
//					System.out.println("======================== End ========================");
				}
			}
		}

		// clone一份出來以免remove時影響到
		ArrayList<Map<String, String>> list2 = (ArrayList) list.clone();
		Iterator<Map<String, String>> iterator = list.iterator();
		Map<String, String> chkMap;
		while (iterator.hasNext()) {
			chkMap = iterator.next();
			for (Map<String, String> chkMap2 : list2) {
				if (chkMap.get("jobEName").equals(chkMap2.get("jobEName"))
						&& chkMap.get("jobPeriod").equals(chkMap2.get("jobPeriod"))
						&& chkMap.get("jobRSDate").equals(chkMap2.get("jobRSDate"))
						&& Long.valueOf(chkMap.get("jobRSOriDateTime")) < Long
								.valueOf(chkMap2.get("jobRSOriDateTime"))) {
					System.out.println("=====list remove=====" + ", jobRSDate = " + chkMap.get("jobRSDate")
							+ ", jobRSTime = " + chkMap.get("jobRSTime") + ", jobRSDateTime = "
							+ chkMap.get("jobRSDateTime") + ", jobRSOriDateTime = " + chkMap.get("jobRSOriDateTime")
							+ ", jobPeriod = " + chkMap.get("jobPeriod") + ", jobSeq = " + chkMap.get("jobSeq")
							+ ", jobEName = " + chkMap.get("jobEName") + ", jobName = " + chkMap.get("jobName")
							+ ", jobRunRS = " + chkMap.get("jobRunRS"));
					iterator.remove();
					break;
				}
			}
		}
	}

	/**
	 * 寫入 "JobList" 頁籤的狀態，並整理出失敗的Job
	 * 
	 * @throws Exception
	 */
	private static void writeSheet1(Sheet sheet1) throws Exception {
		listF = new ArrayList<Map<String, String>>();
		for (Map<String, String> map : list) {
			JobMonth = map.get("jobRSDate").substring(0, 6);
			JobDate = map.get("jobRSDate").substring(6);
			// 判斷是否為當月的日誌
			if (JobMonth.equals(excelMonth)) {

				// 取得對應日期的Cell位置(縱列)
				dateCell = Tools.getDateCell(sheet1, JobDate);

				dataRow = 0;
				// 取得對應Job的Row位置(橫列)
				for (Row row : sheet1) {
					if (Tools.isntBlank(row.getCell(0))
							&& row.getCell(0).getStringCellValue().substring(1).equals(map.get("jobEName"))) {
						dataRow = row.getRowNum();

						// 將執行結果設定至對應位置
						// 因第一行的日欄位有合併儲存格 實際取到的位置為"應檢查"列而非我們要設定的"執行結果"列 故取得的dateCell需+1
						targetRow = sheet1.getRow(dataRow);
						targetCell = targetRow.getCell(dateCell + 1);
						if (dataRow > 0 && (targetCell.getCellType() == Cell.CELL_TYPE_BLANK
								|| map.get("jobRunRS").equals("F"))) {
							targetCell.setCellValue(map.get("jobRunRS"));
							System.out.println("changeCellValue====,dataRow=" + dataRow + ", dateCell=" + dateCell
									+ ", Value=" + map.get("jobRunRS") + ", jobRSDateTime=" + map.get("jobRSDateTime")
									+ ", jobPeriod=" + map.get("jobPeriod"));
							for (Entry<String, String> ent : map.entrySet()) {
								System.out.println(ent.getKey() + " : " + ent.getValue() + " , ");
							}
							// 將失敗的Job放入listF中
							if ("F".equals(map.get("jobRunRS")))
								listF.add(map);
						}
					}
				}
			}else {
				throw new Exception("日誌月份錯誤");
			}
		}

		/**
		 * 寫入最後結果為X的狀態 (前一天為X)
		 * 寫入最後結果為Z的狀態 (今天為檢查的第一天)
		 */
		writeXandZ(sheet1);
	}

	/**
	 * 整理最後結果為X的狀態 (前一天為X)
	 * 整理最後結果為Z的狀態 (今天為檢查的第一天)
	 * 
	 * @throws ParseException @throws
	 */
	private static void writeXandZ(Sheet sheet1) throws ParseException {
		// 檢查的日期(例:20210527)
		int chkDateV = chkDate;
		for (int i = 0; i < Tools.getMinusDays(chkDate); i++) {
			JobMonth = String.valueOf(chkDateV).substring(0, 6);
			if (JobMonth.equals(excelMonth)) {
				JobDate = String.valueOf(chkDateV).substring(6);
				// 取得對應日期的Cell位置(縱列)
				dateCell = Tools.getDateCell(sheet1, JobDate);

				// 取得對應Job的Row位置(橫列)
				for (Row row : sheet1) {
					previouChkCell = row.getCell(dateCell - 2);
					previouCell = row.getCell(dateCell - 1);
					targetChkCell = row.getCell(dateCell);
					targetCell = row.getCell(dateCell + 1);

					// X: 前一天狀態為X，且今天應檢查，且今天尚未壓狀態
					if (previouCell != null && row.getRowNum() > 1 && previouCell.getCellType() == Cell.CELL_TYPE_STRING
							&& previouCell.getStringCellValue().equalsIgnoreCase("X")
							&& targetChkCell.getCellType() == Cell.CELL_TYPE_STRING
							&& targetChkCell.getStringCellValue().equalsIgnoreCase("V")
							&& targetCell.getCellType() == Cell.CELL_TYPE_BLANK) {
						System.out.println("Change Value X : Row=" + row.getRowNum() + ", Cell=" + dateCell);
						targetCell.setCellValue("X");
					}

					// Z: 前一天不應檢查，且今天應檢查，且今天尚未壓狀態
					if (previouChkCell != null && row.getRowNum() > 1
							&& previouChkCell.getCellType() == Cell.CELL_TYPE_BLANK
							&& targetChkCell.getCellType() == Cell.CELL_TYPE_STRING
							&& targetChkCell.getStringCellValue().equalsIgnoreCase("V")
							&& targetCell.getCellType() == Cell.CELL_TYPE_BLANK) {
						System.out.println("Change Value Z : Row=" + row.getRowNum() + ", Cell=" + dateCell);
						targetCell.setCellValue("Z");
					}

					// 需再執行一次有公式的欄位才會更新欄位值
					if (targetCell != null && targetCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						targetCell.setCellFormula(targetCell.getCellFormula());
					}
				}
			}
			chkDateV++;
		}
	}

	/**
	 * 將失敗的job列進 "待辦JOB" 頁籤
	 * 
	 * @throws ParseException
	 */
	private static void writeSheet3(Workbook Workbook) throws ParseException {
		int setColNum = 0, lastRowNum = 0;
		Row row;
		Cell cell = null;
		Map<String, String> map;
		Sheet sheet3 = Workbook.getSheetAt(2);
		Sheet sheet4 = Workbook.getSheetAt(3);
		CellStyle cellStyle = Workbook.createCellStyle();
		lastRowNum = sheet3.getLastRowNum();

		// 為了讓list能由後往前讀，故使用ListIterator
		ListIterator<Map<String, String>> listIterator = listF.listIterator();
		// 先讓迭代器的指標移到最尾筆
		while (listIterator.hasNext()) {
			System.out.println("待辦 job : " + listIterator.next());
		}
		// 再由後往前讀出來
		while (listIterator.hasPrevious()) {
			map = listIterator.previous();
			JobMonth = map.get("jobRSDate").substring(0, 6);
			isPrint = true;
			// 判斷是否為當月的日誌
			if (JobMonth.equals(excelMonth)) {
				// 檢查此項目是否已被列過或者已移至歷史清單
				isPrint = chkSheetForJobF(sheet3, map) && chkSheetForJobF(sheet4, map);

				if (isPrint) {
					lastRowNum++;
					sheet3.createRow(lastRowNum);
					row = sheet3.getRow(lastRowNum);
					String dateStr = map.get("jobRSDate").substring(0, 4) + "/"
							+ map.get("jobRSDate").substring(4, 6) + "/" + map.get("jobRSDate").substring(6);
					// 設定第一欄
					setColNum = 0;
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, dateStr);
					// 設定第二欄
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobSeq"));
					// 設定第三欄
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobEName"));
					// 設定第四欄
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobName"));
					// 設定第五欄
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, map.get("jobPeriod"));
					// 設定第六欄
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, "");
					// 設定第七欄
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, "問題待查");
					// 設定第八欄
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4, "");
				}
			}
		}
	}

	/**
	 * 檢查此項目是否已被列過或者已移至歷史清單
	 * 
	 * @param chkSheet
	 * @param map
	 * @return
	 */
	private static boolean chkSheetForJobF(Sheet chkSheet, Map<String, String> map) {
		for (Row row : chkSheet) {
			targetChkCell = row.getCell(1);
			if (targetChkCell != null && row.getRowNum() > 0 && targetChkCell.getCellType() == Cell.CELL_TYPE_STRING
					&& targetChkCell.getStringCellValue().equals(map.get("jobSeq"))) {
				return false;
			}
		}
		return true;
	}

}
