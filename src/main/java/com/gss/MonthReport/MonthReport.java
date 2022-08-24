package com.gss.MonthReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.gss.Property;

public class MonthReport {

	public static void main(String[] args) {
		String os = System.getProperty("os.name");
		System.out.println("=== NOW TIME ===> " + new Date());
		System.out.println("===os.name===> " + os);

		// 判斷當前執行的啟動方式是IDE還是jar
		boolean isStartupFromJar = new File(
				MaintainList.class.getProtectionDomain().getCodeSource().getLocation().getPath()).isFile();
		System.out.println("isStartupFromJar: " + isStartupFromJar);

		String path = System.getProperty("user.dir") + File.separator; // Jar
		if (!isStartupFromJar) // IDE
			path = os.contains("Mac") ? "/Users/nicole/Dropbox/DailyReport/" // Mac
					: "C:/Users/Nicole/Dropbox/DailyReport/"; // win

		System.out.println("path: " + path);
		Map<String, String> mapProp = Property.getProperties(path);

		// 月報放置路徑與檔名
		String monthReportPath = mapProp.get("MonthReportPath");
		String maintainListExcel = path + monthReportPath+mapProp.get("MaintainListExcel"); // Jar
		String monthReportExcel = path + monthReportPath+mapProp.get("MonthReportExcel"); // Jar
		System.out.println("維護問題紀錄單Excel: " + maintainListExcel);
		System.out.println("月報Excel: " + monthReportExcel);
		
		monthReport(monthReportExcel);
	}

	private static final SimpleDateFormat sdfDateTime = new SimpleDateFormat("yyyy/MM/dd HH:mm");
	private static final SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy/MM/dd");
//	private static final SimpleDateFormat sdfMM = new SimpleDateFormat("yyyyMM");
//	private static final DecimalFormat df = new DecimalFormat("#.#");
//	private static Date acceptDate, replyDate, dueDate, actDate, date;
//	private static String issueType;
//	private static String[] validResultArr;
	private static long diffrence;
	private static final TimeUnit time = TimeUnit.MINUTES;
	private static final List<String> moduleTypeApp = Arrays.asList(new String[] { "ISGS","MD","SD","SC32","CRM","ETL","OTH" });
	private static final List<String> moduleTypeTool = Arrays.asList(new String[] { "SAFENET","DATASTAGE","HADOOP","WEBFOCUS","IA" });
	private static final List<String> listIssueType1 = Arrays.asList(new String[] { "010" });
	private static final List<String> listIssueType2 = Arrays.asList(new String[] { "011", "012" });
	private static final List<String> listIssueType3 = Arrays
			.asList(new String[] { "013", "014", "015", "018", "019" });
//	private static final List<String> listToolsModule = Arrays
//			.asList(new String[] { "SAFENET", "DATASTAGE", "HADOOP", "WEBFOCUS", "IA" });
//	private static List<String> listCell;
//	private static XSSFCellStyle style;
//	private static short dfDecimal, dfNormal, dfDate, dfDateTime;
//	private static Row row;
//
//	private static void setDataFormat(XSSFWorkbook xssfWorkbook) {
//		XSSFDataFormat xssfdf = xssfWorkbook.createDataFormat();
//		dfDecimal = xssfdf.getFormat("#,#0.0"); // 貨幣
//		dfNormal = xssfdf.getFormat(""); // 通用
//		dfDate = xssfdf.getFormat("yyyy/MM/dd");
//		dfDateTime = xssfdf.getFormat("yyyy/MM/dd hh:mm");
//	}
//
//	private static void setVarInit() {
//		listCell = new LinkedList<String>();
//		acceptDate = null;
//		replyDate = null;
//		dueDate = null;
//		actDate = null;
//		date = null;
//		issueType = "";
//		 notFinish = "";
//	}

	public static void monthReport(String monthReportExcel) {
		
		
		 
		XSSFWorkbook xssfWorkbook = null;
		OutputStream output = null;
		
		try {
			File f = new File(monthReportExcel);
			InputStream inputStream = new FileInputStream(f);
			// XSSF (.xlsx)
			xssfWorkbook = new XSSFWorkbook(inputStream);
//			setDataFormat(xssfWorkbook);
			Sheet sheetSummary = xssfWorkbook.getSheetAt(0);
			Sheet sheetStatistics = xssfWorkbook.getSheetAt(1);
//			Sheet sheetMaintainList = xssfWorkbook.getSheetAt(2);
			Sheet sheetPivot = xssfWorkbook.getSheetAt(4);
			
			chkMaintainList(xssfWorkbook.getSheetAt(2));
//			setGoogleDoc(xssfWorkbook);
//			setIA(xssfWorkbook);

			output = new FileOutputStream(f);
			xssfWorkbook.write(output);
		} catch (Exception ex) {
			System.out.println("Error:" + ex.getMessage());
			ex.printStackTrace();
		} finally {
			try {
				if (xssfWorkbook != null)
					xssfWorkbook.close();
				if (output != null)
					output.close();
			} catch (IOException e) {
				System.out.println("Error:" + e.getMessage());
			}
		}
		System.out.println("Done!");
	}

	private static Map<String,String> chkMaintainList(Sheet sheet) throws Exception {
		Date acceptDate, replyDate, dueDate, actDate, date;
		
		int itemCount = 0;
		String issueType = "", validResult = "", notFinish = "", module = "";
		Map<String,String> map = new HashMap<String,String>();
		
		int app010Past = 0, app010Accept = 0, app010Finish = 0, app010NotFinish = 0, app010CodeNum = 0, app010DocNum = 0, app010SA = 0, app010PG = 0 
			, app011Past = 0, app011Accept = 0, app011Finish = 0, app011NotFinish = 0, app011CodeNum = 0, app011DocNum = 0, app011SA = 0, app011PG = 0 
			, app012Past = 0, app012Accept = 0, app012Finish = 0, app012NotFinish = 0, app012CodeNum = 0, app012DocNum = 0, app012SA = 0, app012PG = 0 
			, app013Past = 0, app013Accept = 0, app013Finish = 0, app013NotFinish = 0, app013CodeNum = 0, app013DocNum = 0, app013SA = 0, app013PG = 0 
			, app014Past = 0, app014Accept = 0, app014Finish = 0, app014NotFinish = 0, app014CodeNum = 0, app014DocNum = 0, app014SA = 0, app014PG = 0 
			, app015Past = 0, app015Accept = 0, app015Finish = 0, app015NotFinish = 0, app015CodeNum = 0, app015DocNum = 0, app015SA = 0, app015PG = 0 
			, app018Past = 0, app018Accept = 0, app018Finish = 0, app018NotFinish = 0, app018CodeNum = 0, app018DocNum = 0, app018SA = 0, app018PG = 0 
			, app019Past = 0, app019Accept = 0, app019Finish = 0, app019NotFinish = 0, app019CodeNum = 0, app019DocNum = 0, app019SA = 0, app019PG = 0
			, tool010Past = 0, tool010Accept = 0, tool010Finish = 0, tool010NotFinish = 0, tool010CodeNum = 0, tool010DocNum = 0, tool010SA = 0, tool010PG = 0 
			, tool011Past = 0, tool011Accept = 0, tool011Finish = 0, tool011NotFinish = 0, tool011CodeNum = 0, tool011DocNum = 0, tool011SA = 0, tool011PG = 0 
			, tool012Past = 0, tool012Accept = 0, tool012Finish = 0, tool012NotFinish = 0, tool012CodeNum = 0, tool012DocNum = 0, tool012SA = 0, tool012PG = 0 
			, tool018Past = 0, tool018Accept = 0, tool018Finish = 0, tool018NotFinish = 0, tool018CodeNum = 0, tool018DocNum = 0, tool018SA = 0, tool018PG = 0 
			, tool019Past = 0, tool019Accept = 0, tool019Finish = 0, tool019NotFinish = 0, tool019CodeNum = 0, tool019DocNum = 0, tool019SA = 0, tool019PG = 0;

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet) {
System.out.println(row.getRowNum());
System.out.println(row.getCell(0) != null);
			if (row.getRowNum() >= 5 && row.getCell(0) != null) {

System.out.println(row.getCell(0).getCellFormula());
				if(row.getCell(0).getCellFormula().startsWith("本月"))
					break;

				acceptDate = null;
				replyDate = null;
				dueDate = null;
				actDate = null;
				date = null;
				issueType = "";
				notFinish = "";
				 
				itemCount++;
System.out.println(row.getRowNum());
				// 取Cell資料
				for (int c = 0; c < row.getLastCellNum(); c++) {
					if (row.getCell(c) != null && StringUtils.isNotBlank(row.getCell(c).toString())) {
						
						if (c == 1 || c == 2) {
							date = MonthReportTools.getDateValue(row.getCell(c), sdfDateTime);
							if (c == 1)
								acceptDate = date;
							if (c == 2)
								replyDate = date;
						} else if (c == 5) {
							module = row.getCell(c).toString();
							if(moduleTypeApp.contains(module.toUpperCase()))
								module = "app";
							else
								module = "tool";
						} else if (c == 7) {
							issueType = row.getCell(c).toString();
						} else if (c == 9 || c == 10) {
							date = MonthReportTools.getDateValue(row.getCell(c), sdfDate);
							if (c == 9)
								dueDate = date;
							if (c == 10)
								actDate = date;
						} else if ((c >= 11 && c <= 14) || c == 16 || c == 17) {
							notFinish += row.getCell(c).toString();
						}
					}

				}
				
				// ============================ 計算彙總筆數 =======================================
System.out.println(
"\r\napp010Past = "+app010Past+", app010Accept = "+app010Accept+", app010Finish = "+app010Finish+", app010NotFinish = "+app010NotFinish+", app010CodeNum = "+app010CodeNum+", app010DocNum = "+app010DocNum+", app010SA = "+app010SA+", app010PG = "+app010PG
+"\r\n, app011Past = "+app011Past+", app011Accept = "+app011Accept+", app011Finish = "+app011Finish+", app011NotFinish = "+app011NotFinish+", app011CodeNum = "+app011CodeNum+", app011DocNum = "+app011DocNum+", app011SA = "+app011SA+", app011PG = "+app011PG
+"\r\n, app012Past = "+app012Past+", app012Accept = "+app012Accept+", app012Finish = "+app012Finish+", app012NotFinish = "+app012NotFinish+", app012CodeNum = "+app012CodeNum+", app012DocNum = "+app012DocNum+", app012SA = "+app012SA+", app012PG = "+app012PG
+"\r\n, app013Past = "+app013Past+", app013Accept = "+app013Accept+", app013Finish = "+app013Finish+", app013NotFinish = "+app013NotFinish+", app013CodeNum = "+app013CodeNum+", app013DocNum = "+app013DocNum+", app013SA = "+app013SA+", app013PG = "+app013PG
+"\r\n, app014Past = "+app014Past+", app014Accept = "+app014Accept+", app014Finish = "+app014Finish+", app014NotFinish = "+app014NotFinish+", app014CodeNum = "+app014CodeNum+", app014DocNum = "+app014DocNum+", app014SA = "+app014SA+", app014PG = "+app014PG
+"\r\n, app015Past = "+app015Past+", app015Accept = "+app015Accept+", app015Finish = "+app015Finish+", app015NotFinish = "+app015NotFinish+", app015CodeNum = "+app015CodeNum+", app015DocNum = "+app015DocNum+", app015SA = "+app015SA+", app015PG = "+app015PG
+"\r\n, app018Past = "+app018Past+", app018Accept = "+app018Accept+", app018Finish = "+app018Finish+", app018NotFinish = "+app018NotFinish+", app018CodeNum = "+app018CodeNum+", app018DocNum = "+app018DocNum+", app018SA = "+app018SA+", app018PG = "+app018PG
+"\r\n, app019Past = "+app019Past+", app019Accept = "+app019Accept+", app019Finish = "+app019Finish+", app019NotFinish = "+app019NotFinish+", app019CodeNum = "+app019CodeNum+", app019DocNum = "+app019DocNum+", app019SA = "+app019SA+", app019PG = "+app019PG
+"\r\n, tool010Past = "+tool010Past+", tool010Accept = "+tool010Accept+", tool010Finish = "+tool010Finish+", tool010NotFinish = "+tool010NotFinish+", tool010CodeNum = "+tool010CodeNum+", tool010DocNum = "+tool010DocNum+", tool010SA = "+tool010SA+", tool010PG = "+tool010PG
+"\r\n, tool011Past = "+tool011Past+", tool011Accept = "+tool011Accept+", tool011Finish = "+tool011Finish+", tool011NotFinish = "+tool011NotFinish+", tool011CodeNum = "+tool011CodeNum+", tool011DocNum = "+tool011DocNum+", tool011SA = "+tool011SA+", tool011PG = "+tool011PG
+"\r\n, tool012Past = "+tool012Past+", tool012Accept = "+tool012Accept+", tool012Finish = "+tool012Finish+", tool012NotFinish = "+tool012NotFinish+", tool012CodeNum = "+tool012CodeNum+", tool012DocNum = "+tool012DocNum+", tool012SA = "+tool012SA+", tool012PG = "+tool012PG
+"\r\n, tool018Past = "+tool018Past+", tool018Accept = "+tool018Accept+", tool018Finish = "+tool018Finish+", tool018NotFinish = "+tool018NotFinish+", tool018CodeNum = "+tool018CodeNum+", tool018DocNum = "+tool018DocNum+", tool018SA = "+tool018SA+", tool018PG = "+tool018PG
+"\r\n, tool019Past = "+tool019Past+", tool019Accept = "+tool019Accept+", tool019Finish = "+tool019Finish+", tool019NotFinish = "+tool019NotFinish+", tool019CodeNum = "+tool019CodeNum+", tool019DocNum = "+tool019DocNum+", tool019SA = "+tool019SA+", tool019PG = "+tool019PG);
				
				// 過去的單
				if(MonthReportTools.isPastDate(acceptDate)) {
					// 應用系統類別
					if ("app".equals(module)) {
						// 問題類型
						if ("010".equals(issueType)) {
							app010Past++;
							if (actDate == null)
								app010NotFinish++;
							else
								app010Finish++;
						} else if ("011".equals(issueType)) {
							app011Past++;
							if (actDate == null)
								app011NotFinish++;
							else
								app011Finish++;
						} else if ("012".equals(issueType)) {
							app012Past++;
							if (actDate == null)
								app012NotFinish++;
							else
								app012Finish++;
						} else if ("013".equals(issueType)) {
							app013Past++;
							if (actDate == null)
								app013NotFinish++;
							else
								app013Finish++;
						} else if ("014".equals(issueType)) {
							app014Past++;
							if (actDate == null)
								app014NotFinish++;
							else
								app014Finish++;
						} else if ("015".equals(issueType)) {
							app015Past++;
							if (actDate == null)
								app015NotFinish++;
							else
								app015Finish++;
						} else if ("018".equals(issueType)) {
							app018Past++;
							if (actDate == null)
								app018NotFinish++;
							else
								app018Finish++;
						} else if ("019".equals(issueType)) {
							app019Past++;
							if (actDate == null)
								app019NotFinish++;
							else
								app019Finish++;
						}
					} else { // 工具軟體類別
						if ("010".equals(issueType)) {
							tool010Past++;
							if (actDate == null)
								tool010NotFinish++;
							else
								tool010Finish++;
						} else if ("011".equals(issueType)) {
							tool011Past++;
							if (actDate == null)
								tool011NotFinish++;
							else
								tool011Finish++;
						} else if ("012".equals(issueType)) {
							tool012Past++;
							if (actDate == null)
								tool012NotFinish++;
							else
								tool012Finish++;
						} else if ("018".equals(issueType)) {
							tool018Past++;
							if (actDate == null)
								tool018NotFinish++;
							else
								tool018Finish++;
						} else if ("019".equals(issueType)) {
							tool019Past++;
							if (actDate == null)
								tool019NotFinish++;
							else
								tool019Finish++;
						}
					}
				} else { // 當月的單
					// 應用系統類別
					if ("app".equals(module)) {
						// 問題類型
						if ("010".equals(issueType)) {
							app010Accept++;
							if (actDate == null)
								app010NotFinish++;
							else
								app010Finish++;
						} else if ("011".equals(issueType)) {
							app011Accept++;
							if (actDate == null)
								app011NotFinish++;
							else
								app011Finish++;
						} else if ("012".equals(issueType)) {
							app012Accept++;
							if (actDate == null)
								app012NotFinish++;
							else
								app012Finish++;
						} else if ("013".equals(issueType)) {
							app013Accept++;
							if (actDate == null)
								app013NotFinish++;
							else
								app013Finish++;
						} else if ("014".equals(issueType)) {
							app014Accept++;
							if (actDate == null)
								app014NotFinish++;
							else
								app014Finish++;
						} else if ("015".equals(issueType)) {
							app015Accept++;
							if (actDate == null)
								app015NotFinish++;
							else
								app015Finish++;
						} else if ("018".equals(issueType)) {
							app018Accept++;
							if (actDate == null)
								app018NotFinish++;
							else
								app018Finish++;
						} else if ("019".equals(issueType)) {
							app019Accept++;
							if (actDate == null)
								app019NotFinish++;
							else
								app019Finish++;
						}
					} else { // 工具軟體類別
						if ("010".equals(issueType)) {
							tool010Accept++;
							if (actDate == null)
								tool010NotFinish++;
							else
								tool010Finish++;
						} else if ("011".equals(issueType)) {
							tool011Accept++;
							if (actDate == null)
								tool011NotFinish++;
							else
								tool011Finish++;
						} else if ("012".equals(issueType)) {
							tool012Accept++;
							if (actDate == null)
								tool012NotFinish++;
							else
								tool012Finish++;
						} else if ("018".equals(issueType)) {
							tool018Accept++;
							if (actDate == null)
								tool018NotFinish++;
							else
								tool018Finish++;
						} else if ("019".equals(issueType)) {
							tool019Accept++;
							if (actDate == null)
								tool019NotFinish++;
							else
								tool019Finish++;
						}
					}
				}
				// ============================ 計算彙總筆數 END =======================================
					
				if ((actDate == null || (actDate != null && MonthReportTools.isFutureDate(actDate)))
						&& !"0.00.0".equals(notFinish))
					throw new Exception("notFinish後面欄位未清空");
				
				validResult = MonthReportTools.getValidResult(acceptDate, replyDate, dueDate, actDate, issueType, "");
				if(validResult.startsWith("ERR"))
					throw new Exception(validResult);
			}
		}
		
		map.put("itemCount", String.valueOf(itemCount));
		return map;

	}

}
