package com.gss;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.ListIterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class RunDailyReport {
	static final String mailStartText = " - 您好, 〔("; // job 寄的 mail 開頭
	static Integer dateCell = 0, dataRow = 0, chkDate = 0, runtimeInt = 0;
	static String JobMonth = "", JobDate = "", excelMonth = "", DailyReportExcel = "", account = "", pwd = "",
			runtime = "";
	static boolean isPrint;
	static Row targetRow;
	static Cell targetCell, previouCell, targetChkCell, previouChkCell, runtimeCell;
	static ArrayList<Map<String, String>> listF, list;
	static ArrayList<TreeMap<String, String>> listFforSheet3;
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
		chkDate = "auto".equals(mapProp.get("chkDate")) ? Tools.getChkDate() : Integer.parseInt(mapProp.get("chkDate"));
		System.out.println("要從 " + chkDate + " 的mail開始取 ");

		// 日報放置路徑與檔名
		DailyReportExcel = path + mapProp.get("DailyReportExcel"); // Jar
		System.out.println("日誌Excel: " + DailyReportExcel);

		// 收件匣名稱
		inboxName = mapProp.get("inboxName").split(",");
		String inboxNameStr = "";
		for (String str : inboxName)
			inboxNameStr += inboxNameStr.length() > 0 ? ", " + str : str;
		System.out.println("收件匣名稱: " + inboxNameStr);

		// 帳號
		account = mapProp.get("account");
//		System.out.println("Mail帳號: " + account);

		// 密碼
		pwd = mapProp.get("pwd");
//		System.out.println("Mail密碼: " + pwd);

		Workbook workbook = null;
		OutputStream output = null;

		try {
			// 整理 MAIL內容
			parserMailContent(path);

			File f = new File(DailyReportExcel);
			workbook = Tools.getWorkbook(DailyReportExcel, f);
			Sheet sheet1 = workbook.getSheetAt(0);

			// 日誌的月份 年月(六碼)
			excelMonth = sheet1.getRow(0).getCell(0).getStringCellValue().trim();
			excelMonth = excelMonth.substring(0, 7).trim();

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

		// 將失敗的job寫入file中 (填寫日誌清單_2021)
		String txt = "";
		for (Map<String, String> map : listF)
			if (!txt.contains(map.get("jobName")))
				txt += map.get("jobName") + "\r\n";
		Tools.writeListFtoFile(path, "填寫日誌清單用 \r\n\r\n" + txt, false);
	}

	/**
	 * 整理 MAIL內容
	 * 
	 * @return 整理後的list
	 *         List<Map<String, String>>
	 *         map.put("jobRSDate", jobRSDate); // 日誌日期 YYYYMMDD
	 *         map.put("jobRSTime", jobRSTime); // job時間 HHmm
	 *         map.put("jobRSDateTime", jobRSDate + " " + jobRSTime); // 日誌日期時間 YYYYMMDD HHmm
	 *         map.put("jobRSOriDateTime", jobRSOriDate + " " + jobRSTime); // job原日期時間 (後面比對會用到)
	 *         map.put("jobPeriod", jobPeriod); // job執行區間
	 *         map.put("jobSeq", jobSeq); // jobSeq
	 *         map.put("jobEName", jobEName); // 英文名
	 *         map.put("jobName", jobName); // 中文名
	 *         map.put("jobRunRS", jobRunRS); // 執行結果
	 * @throws ParseException
	 */
	private static void parserMailContent(String path) throws ParseException {
		boolean isPm = false;
		int hhInt = 0, arrLen = 0;
		String jobRSText = "", jobRSDate = "", jobRSOriDate = "", jobRSTime = "", jobPeriod = "",
				jobSeq = "", jobEName = "", jobName = "", jobRunRS = "", time = "";
		String[] jobMailTitleArr;
		Map<String, String> map;
		list = new ArrayList<Map<String, String>>();
		List<Map<String, String>> listMail;
		listFforSheet3 = new ArrayList<TreeMap<String, String>>();
		
		/**
		 *  取 昨、今 兩天的日期
		 *  因mail中的日期會寫"昨天"、"今天"而非日期
		 */
		Calendar cal = Calendar.getInstance();
		String today = Tools.getCalendar2String(cal, "yyyyMMdd");
		cal.add(Calendar.DATE, -1);
		String yesterday = Tools.getCalendar2String(cal, "yyyyMMdd");

		// listMail
		cal.setTime(Tools.getString2Date(String.valueOf(chkDate), "yyyyMMdd"));
		listMail = Selenium_Crawler.getMailContent(path, inboxName, account, pwd, cal, listFforSheet3);

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
				time = jobMailTitleArr[arrLen - 1].trim(); // 下午10:22
				isPm = time.substring(0, 2).equals("下午"); // 上午false 下午true
				time = time.substring(2); // 10:22
				hhInt = Integer.parseInt(time.substring(0, time.indexOf(":"))); // 10
				hhInt = hhInt == 12 ? 0 : hhInt; // 0 ~ 11
				hhInt = hhInt + (isPm ? 12 : 0); // 0 ~ 23
				// mail收到的時間(24小時制)
				jobRSTime = Tools.getLen2(String.valueOf(hhInt)) + time.substring(time.indexOf(":")); // 22:22
				// job 所屬日期 (日誌用)
				jobRSDate = jobMailTitleArr[arrLen - 2].trim();
				// job 原日期 (後面刪除相同job時使用)
				jobRSOriDate = jobRSDate.equals("昨天") ? yesterday
						: jobRSDate.equals("今天") ? today
								: Tools.getDate2String(
										Tools.getString2Date(jobRSDate.substring(0, jobRSDate.length() - 1), "yy/M/d"),
										"yyyyMMdd");
				System.out.println("=== jobRSDate : " + jobRSDate + " , jobRSOriDate : " + jobRSOriDate
						+ " , jobRSTime : " + jobRSTime);
				cal.setTime(Tools.getString2Date(jobRSOriDate + " " + jobRSTime, "yyyyMMdd HH:mm"));
				System.out.println("===== jobRSOriDate YYYYMMDDHHmm ====> " + Tools.getCalendar2String(cal, "yyyyMMdd HH:mm"));
				
				// job 執行區間
				jobPeriod = jobMailTitleArr[arrLen - 5]; // 因job中文名稱內可能會有多個逗號，故抓倒數第五位陣列值
				jobPeriod = jobPeriod.substring(jobPeriod.lastIndexOf("_") + 1, jobPeriod.lastIndexOf("("));
				if (hhInt > 8)
					cal.add(Calendar.DATE, 1);
				// job 所屬日期 (日誌用)
				jobRSDate = Tools.getCalendar2String(cal, "yyyyMMdd");

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
					map.put("jobRSDateTime", jobRSDate + " " + jobRSTime);
					map.put("jobRSOriDateTime", jobRSOriDate + " " + jobRSTime);
					map.put("jobPeriod", jobPeriod);
					map.put("jobSeq", jobSeq);
					map.put("jobEName", jobEName);
					map.put("jobName", jobName);
					map.put("jobRunRS", jobRunRS);
					list.add(map);

					System.out.println("======================== Start ========================");
//					System.out.println("jobRSDate 日誌日期 => " + jobRSDate);
//					System.out.println("jobRSTime job時間 => " + jobRSTime);
					System.out.println("jobRSDateTime 日誌日期時間 => " + jobRSDate + " " + jobRSTime);
					System.out.println("jobRSOriDateTime job原日期時間 => " + jobRSOriDate + " " + jobRSTime);
					System.out.println("jobPeriod job執行區間 => " + jobPeriod);
					System.out.println("jobSeq => " + jobSeq);
					System.out.println("jobEName 英文名=> " + jobEName);
					System.out.println("jobName 中文名 => " + jobName);
					System.out.println("jobRunRS 執行結果 => " + jobRunRS);
					System.out.println("======================== End ========================");
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
						&& Long.valueOf(chkMap.get("jobRSOriDateTime").replaceAll("[ :]", "")) < Long
								.valueOf(chkMap2.get("jobRSOriDateTime").replaceAll("[ :]", ""))) {
					System.out.println("=====list remove===== "
							+ "jobRSDateTime = " + chkMap.get("jobRSDateTime") + ", jobRSOriDateTime = " + chkMap.get("jobRSOriDateTime")
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
						&& row.getCell(0).getStringCellValue().replaceAll("\u00A0", "").equals(map.get("jobEName"))
					) {
						dataRow = row.getRowNum();

						// 將執行結果設定至對應位置
						// 因第一行的日欄位有合併儲存格 實際取到的位置為"應檢查"列而非我們要設定的"執行結果"列 故取得的dateCell需+1
						targetRow = sheet1.getRow(dataRow);
						targetCell = targetRow.getCell(dateCell + 1);
						// 應檢查
						if ("V".equals(targetRow.getCell(dateCell).getStringCellValue())) {
							/**
							 * 成功的job
							 * 寫入Sheet1的狀態(當targetCell不為F時則壓S)
							 */
							if (map.get("jobRunRS").equals("S")
									&& (targetCell.getCellType() == Cell.CELL_TYPE_BLANK
											|| (targetCell.getCellType() == Cell.CELL_TYPE_STRING
													&& !"F".equals(targetCell.getStringCellValue()))))
								targetCell.setCellValue(map.get("jobRunRS"));

							/**
							 * 失敗的job
							 * 寫入Sheet1的狀態
							 * 寫入JobF.txt
							 */
							if ("F".equals(map.get("jobRunRS"))) {
								targetCell.setCellValue(map.get("jobRunRS"));
								listF.add(map);
							}
						}
						// 將失敗的Job放入job待辦頁籤中 (不論是否應檢查)
						if ("F".equals(map.get("jobRunRS"))) {
							for(TreeMap<String,String> mapRQ : listFforSheet3) {
								if(mapRQ.get("RQ_job_seq").equals(map.get("jobSeq"))){
									mapRQ.putAll(map);
									System.out.println("==== map putAll to mapRQ ===="); 
									for (Entry<String, String> ent : mapRQ.entrySet()) {
										System.out.println(ent.getKey() + " : " + ent.getValue() + " , ");
									}
								}
							}
						}

						System.out.println("changeCellValue====> dataRow=" + dataRow + ", dateCell=" + dateCell
								+ ", Value=" + map.get("jobRunRS") + ", jobRSDateTime=" + map.get("jobRSDateTime")
								+ ", jobPeriod=" + map.get("jobPeriod"));

						for (Entry<String, String> ent : map.entrySet()) {
							System.out.println(ent.getKey() + " : " + ent.getValue() + " , ");
						}
					}
				}
			}else {
				System.out.println("日誌月份錯誤");
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
					previouChkCell = row.getCell(dateCell - 2); //前一天的"應檢查"
					previouCell = row.getCell(dateCell - 1); //前一天的"執行結果"
					targetChkCell = row.getCell(dateCell); //今天的"應檢查"
					targetCell = row.getCell(dateCell + 1); //今天的"執行結果"
					runtimeCell = row.getCell(6); //執行時間

					/**
					 * X: 前一天狀態為X && 今天應檢查 && 今天尚未壓狀態
					 * 月初不會有X
					 */
					if (previouCell != null && row.getRowNum() > 1 
							&& previouCell.getCellType() == Cell.CELL_TYPE_STRING
							&& previouCell.getStringCellValue().equalsIgnoreCase("X")
							&& targetChkCell.getCellType() == Cell.CELL_TYPE_STRING
							&& targetChkCell.getStringCellValue().equalsIgnoreCase("V")
							&& targetCell.getCellType() == Cell.CELL_TYPE_BLANK) {
						System.out.println("Change Value X : Row=" + row.getRowNum() + ", Cell=" + dateCell);
						targetCell.setCellValue("X");
					}

					/**
					 * Z: 前一天不應檢查 && 今天應檢查 && 今天尚未壓狀態 && 執行時間 >= 09:00
					 * 含月初的判斷
					 */
					if (previouChkCell != null && row.getRowNum() > 1
							&& (previouChkCell.getCellType() == Cell.CELL_TYPE_BLANK
									|| (previouChkCell.getCellType() == Cell.CELL_TYPE_STRING
											&& !previouChkCell.getStringCellValue().equalsIgnoreCase("V")))
							&& targetChkCell.getCellType() == Cell.CELL_TYPE_STRING
							&& targetChkCell.getStringCellValue().equalsIgnoreCase("V")
							&& targetCell.getCellType() == Cell.CELL_TYPE_BLANK) {
						runtime = runtimeCell.getStringCellValue().replaceAll("\u00A0", "");
						runtimeInt = Integer.parseInt(runtime.substring(0, runtime.indexOf(":")));
						if (runtimeInt >= 9) {
							System.out.println("Change Value Z : Row=" + row.getRowNum() + ", Cell=" + dateCell + ", runtime=" + runtime);
							targetCell.setCellValue("Z");
						}
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
		ListIterator<TreeMap<String, String>> listIterator = listFforSheet3.listIterator();
		// 先讓迭代器的指標移到最尾筆
		while (listIterator.hasNext()) {
			System.out.println("待辦 job : " + listIterator.next());
		}
		// 再由後往前讀出來
		while (listIterator.hasPrevious()) {
			map = listIterator.previous();
			// 若出現不在當月日誌清單內的失敗job則跳過
			if(map.get("jobRSDate") == null) {
				continue;
			}
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
					Tools.setCellStyle(setColNum++, cell, cellStyle, row, sheet3, sheet4,
							map.get("RQ_rq_id") + " => " + map.get("RQ_run_flag"));
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
