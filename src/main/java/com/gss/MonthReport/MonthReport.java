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
			, app015Past = 0, app015Accept = 0, app0150Finish = 0, app015NotFinish = 0, app015CodeNum = 0, app015DocNum = 0, app015SA = 0, app015PG = 0 
			, app018Past = 0, app018Accept = 0, app018Finish = 0, app018NotFinish = 0, app018CodeNum = 0, app018DocNum = 0, app018SA = 0, app018PG = 0 
			, app019Past = 0, app019Accept = 0, app019Finish = 0, app019NotFinish = 0, app019CodeNum = 0, app019DocNum = 0, app019SA = 0, app019PG = 0
			, tool010Past = 0, tool010Accept = 0, tool010Finish = 0, tool010NotFinish = 0, tool010CodeNum = 0, tool010DocNum = 0, tool010SA = 0, tool010PG = 0 
			, tool011Past = 0, tool011Accept = 0, tool011Finish = 0, tool011NotFinish = 0, tool011CodeNum = 0, tool011DocNum = 0, tool011SA = 0, tool011PG = 0 
			, tool012Past = 0, tool012Accept = 0, tool012Finish = 0, tool012NotFinish = 0, tool012CodeNum = 0, tool012DocNum = 0, tool012SA = 0, tool012PG = 0 
			, tool018Past = 0, tool018Accept = 0, tool018Finish = 0, tool018NotFinish = 0, tool018CodeNum = 0, tool018DocNum = 0, tool018SA = 0, tool018PG = 0 
			, tool019Past = 0, tool019Accept = 0, tool019Finish = 0, tool019NotFinish = 0, tool019CodeNum = 0, tool019DocNum = 0, tool019SA = 0, tool019PG = 0;

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet) {
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
				
				if(MonthReportTools.isPastDate(acceptDate)) {
					if ("app".equals(module)) {
						if ("010".equals(issueType))
							app010Past++;
						if ("011".equals(issueType))
							app011Past++;
						if ("012".equals(issueType))
							app012Past++;
						if ("013".equals(issueType))
							app013Past++;
						if ("014".equals(issueType))
							app014Past++;
						if ("015".equals(issueType))
							app015Past++;
						if ("018".equals(issueType))
							app018Past++;
						if ("019".equals(issueType))
							app019Past++;
					} else {
						if ("010".equals(issueType))
							tool010Past++;
						if ("011".equals(issueType))
							tool011Past++;
						if ("012".equals(issueType))
							tool012Past++;
						if ("018".equals(issueType))
							tool018Past++;
						if ("019".equals(issueType))
							tool019Past++;
					}
				}
					
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
