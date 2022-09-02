package com.gss.MonthReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
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
		int lastMonthTotalFinishNum = 228;
		int lastMonthTotalCodeNum = 178;
		monthReport(monthReportExcel, lastMonthTotalFinishNum, lastMonthTotalCodeNum);
	}

	private static final SimpleDateFormat sdfDateTime = new SimpleDateFormat("yyyy/MM/dd HH:mm");
	private static final SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy/MM/dd");
	private static final List<String> moduleTypeApp = Arrays.asList(new String[] { "ISGS","MD","SD","SC32","CRM","ETL","OTH" });

	public static void monthReport(String monthReportExcel, Integer lastMonthTotalFinishNum, Integer lastMonthTotalCodeNum) {
		 
		XSSFWorkbook xssfWorkbook = null;
		
		try {
			File f = new File(monthReportExcel);
			InputStream inputStream = new FileInputStream(f);
			// XSSF (.xlsx)
			xssfWorkbook = new XSSFWorkbook(inputStream);
			
			Map<String, Float> maintainRS = chkMaintainList("系統維護問題紀錄表", xssfWorkbook.getSheetAt(2), lastMonthTotalFinishNum, lastMonthTotalCodeNum);
			chkSummaryList("彙總表", xssfWorkbook.getSheetAt(0), maintainRS);
			chkSummaryMonthList("月統計", xssfWorkbook.getSheetAt(1), maintainRS);
			chkPIVOT("PIVOT", xssfWorkbook.getSheetAt(4), maintainRS);

		} catch (Exception ex) {
			System.out.println("Error:" + ex.getMessage());
			ex.printStackTrace();
		} finally {
			try {
				if (xssfWorkbook != null)
					xssfWorkbook.close();
			} catch (IOException e) {
				System.out.println("Error:" + e.getMessage());
			}
		}
		System.out.println("Done!");
	}


	private static void chkPIVOT(String sheetName, Sheet sheet, Map<String, Float> maintainRS) throws Exception {
		
		if(sheet.getRow(6).getCell(1).getNumericCellValue() != (Double.parseDouble(maintainRS.get("appCountFinish").toString()) + Double.parseDouble(maintainRS.get("appCountNotFinish").toString()))) throw new Exception(sheetName + " Y 計數 - 紀錄單號 數值錯誤!");
		if(sheet.getRow(6).getCell(2).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountCodeNum").toString())) throw new Exception(sheetName + " Y 加總 - 程式支數 數值錯誤!");
		if(sheet.getRow(6).getCell(3).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountDocNum").toString())) throw new Exception(sheetName + " Y 加總 - 文件數 數值錯誤!");
		if(sheet.getRow(6).getCell(4).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountSA").toString())) throw new Exception(sheetName + " Y 加總 - 系統分析人時 數值錯誤!");
		if(sheet.getRow(6).getCell(5).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountPG").toString())) throw new Exception(sheetName + " Y 加總 - 程式設計人時 數值錯誤!");

		if(sheet.getRow(13).getCell(1).getNumericCellValue() != (Double.parseDouble(maintainRS.get("appCountFinish").toString()) + Double.parseDouble(maintainRS.get("appCountNotFinish").toString()))) throw new Exception(sheetName + " 總計 計數 - 紀錄單號 數值錯誤!");
		if(sheet.getRow(13).getCell(2).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountCodeNum").toString())) throw new Exception(sheetName + " 總計 加總 - 程式支數 數值錯誤!");
		if(sheet.getRow(13).getCell(3).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountDocNum").toString())) throw new Exception(sheetName + " 總計 加總 - 文件數 數值錯誤!");
		if(sheet.getRow(13).getCell(4).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountSA").toString())) throw new Exception(sheetName + " 總計 加總 - 系統分析人時 數值錯誤!");
		if(sheet.getRow(13).getCell(5).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountPG").toString())) throw new Exception(sheetName + " 總計 加總 - 程式設計人時 數值錯誤!");
	}
	
	private static void chkSummaryMonthList(String sheetName, Sheet sheet, Map<String, Float> maintainRS) throws Exception {
		int lastMonth = MonthReportTools.getLastMonth();
		int rowNum = lastMonth + 4;
		
		// 驗證上半部細項數值
		if(sheet.getRow(rowNum).getCell(1).getNumericCellValue() != (sheet.getRow(rowNum).getCell(3).getNumericCellValue() + sheet.getRow(rowNum).getCell(4).getNumericCellValue())) throw new Exception(sheetName + " 服務次數 數值錯誤!");
		if(sheet.getRow(rowNum).getCell(2).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountAccept").toString())) throw new Exception(sheetName + " 受理件數 數值錯誤!");
		if(sheet.getRow(rowNum).getCell(3).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountFinish").toString())) throw new Exception(sheetName + " 完成件數 數值錯誤!");
		if(sheet.getRow(rowNum).getCell(4).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountNotFinish").toString())) throw new Exception(sheetName + " 未完成件數 數值錯誤!");
		if(sheet.getRow(rowNum).getCell(9).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountCodeNum").toString())) throw new Exception(sheetName + " 程式支數 數值錯誤!");
		if(sheet.getRow(rowNum).getCell(10).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountDocNum").toString())) throw new Exception(sheetName + " 文件數 數值錯誤!");
		if(sheet.getRow(rowNum).getCell(11).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountSA").toString())) throw new Exception(sheetName + " 系統分析人時 數值錯誤!");
		if(sheet.getRow(rowNum).getCell(12).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountPG").toString())) throw new Exception(sheetName + " 程式設計人時 數值錯誤!");

		// 驗證下半部數值
		if(sheet.getRow(18).getCell(11).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountSA").toString())) throw new Exception(sheetName + " 系統分析人時 已使用時數 (A) 數值錯誤!");
		if(sheet.getRow(18).getCell(12).getNumericCellValue() != Double.parseDouble(maintainRS.get("appCountPG").toString())) throw new Exception(sheetName + " 程式設計人時 已使用時數 (A) 數值錯誤!");
		
		// 驗證下半部月份
		if(Integer.parseInt(sheet.getRow(18).getCell(9).getStringCellValue().split("/")[1].split(" ")[0]) != lastMonth) throw new Exception(sheetName + " 已使用時數 (A) 月份 錯誤!");
		if(Integer.parseInt(sheet.getRow(20).getCell(9).getStringCellValue().split("/")[1].split(" ")[0]) != lastMonth) throw new Exception(sheetName + " 可核銷金額 (A*B) 月份 錯誤!");
		if(Integer.parseInt(sheet.getRow(21).getCell(9).getStringCellValue().split("/")[1].split(" ")[0]) != lastMonth) throw new Exception(sheetName + " 核銷金額總計 月份 錯誤!");
		if(Integer.parseInt(sheet.getRow(24).getCell(9).getStringCellValue().split("/")[1].split(" ")[0]) != lastMonth) throw new Exception(sheetName + " 已使用時數 (D) 月份 錯誤!");

	}
	

	private static void chkSummaryList(String sheetName, Sheet sheet, Map<String, Float> maintainRS) throws Exception {

		// 驗證彙總表的日期區間
		String period = sheet.getRow(2).getCell(0).getStringCellValue();
		String[] periodArr = period.split("/");
		if(MonthReportTools.getLastMonth() != Integer.parseInt(periodArr[1]) || MonthReportTools.getLastMonth() != Integer.parseInt(periodArr[3])) throw new Exception(sheetName + " 月份錯誤!");
		
		// 驗證彙總表的各欄位內容
		if(Double.parseDouble(maintainRS.get("app010Past").toString()) != sheet.getRow(6).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app010Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app010Accept").toString()) != sheet.getRow(6).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app010Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app010Finish").toString()) != sheet.getRow(6).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app010Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app010NotFinish").toString()) != sheet.getRow(6).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app010NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app010CodeNum").toString()) != sheet.getRow(6).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app010CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app010DocNum").toString()) != sheet.getRow(6).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app010DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app010SA").toString()) != sheet.getRow(6).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app010SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app010PG").toString()) != sheet.getRow(6).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app010PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("app011Past").toString()) != sheet.getRow(8).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app011Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app011Accept").toString()) != sheet.getRow(8).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app011Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app011Finish").toString()) != sheet.getRow(8).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app011Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app011NotFinish").toString()) != sheet.getRow(8).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app011NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app011CodeNum").toString()) != sheet.getRow(8).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app011CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app011DocNum").toString()) != sheet.getRow(8).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app011DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app011SA").toString()) != sheet.getRow(8).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app011SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app011PG").toString()) != sheet.getRow(8).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app011PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("app012Past").toString()) != sheet.getRow(9).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app012Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app012Accept").toString()) != sheet.getRow(9).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app012Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app012Finish").toString()) != sheet.getRow(9).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app012Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app012NotFinish").toString()) != sheet.getRow(9).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app012NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app012CodeNum").toString()) != sheet.getRow(9).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app012CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app012DocNum").toString()) != sheet.getRow(9).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app012DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app012SA").toString()) != sheet.getRow(9).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app012SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app012PG").toString()) != sheet.getRow(9).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app012PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("app013Past").toString()) != sheet.getRow(11).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app013Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app013Accept").toString()) != sheet.getRow(11).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app013Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app013Finish").toString()) != sheet.getRow(11).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app013Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app013NotFinish").toString()) != sheet.getRow(11).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app013NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app013CodeNum").toString()) != sheet.getRow(11).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app013CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app013DocNum").toString()) != sheet.getRow(11).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app013DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app013SA").toString()) != sheet.getRow(11).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app013SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app013PG").toString()) != sheet.getRow(11).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app013PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("app014Past").toString()) != sheet.getRow(12).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app014Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app014Accept").toString()) != sheet.getRow(12).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app014Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app014Finish").toString()) != sheet.getRow(12).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app014Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app014NotFinish").toString()) != sheet.getRow(12).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app014NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app014CodeNum").toString()) != sheet.getRow(12).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app014CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app014DocNum").toString()) != sheet.getRow(12).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app014DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app014SA").toString()) != sheet.getRow(12).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app014SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app014PG").toString()) != sheet.getRow(12).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app014PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("app015Past").toString()) != sheet.getRow(13).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app015Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app015Accept").toString()) != sheet.getRow(13).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app015Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app015Finish").toString()) != sheet.getRow(13).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app015Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app015NotFinish").toString()) != sheet.getRow(13).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app015NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app015CodeNum").toString()) != sheet.getRow(13).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app015CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app015DocNum").toString()) != sheet.getRow(13).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app015DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app015SA").toString()) != sheet.getRow(13).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app015SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app015PG").toString()) != sheet.getRow(13).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app015PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("app018Past").toString()) != sheet.getRow(14).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app018Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app018Accept").toString()) != sheet.getRow(14).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app018Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app018Finish").toString()) != sheet.getRow(14).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app018Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app018NotFinish").toString()) != sheet.getRow(14).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app018NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app018CodeNum").toString()) != sheet.getRow(14).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app018CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app018DocNum").toString()) != sheet.getRow(14).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app018DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app018SA").toString()) != sheet.getRow(14).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app018SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app018PG").toString()) != sheet.getRow(14).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app018PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("app019Past").toString()) != sheet.getRow(15).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " app019Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app019Accept").toString()) != sheet.getRow(15).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " app019Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app019Finish").toString()) != sheet.getRow(15).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " app019Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app019NotFinish").toString()) != sheet.getRow(15).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " app019NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app019CodeNum").toString()) != sheet.getRow(15).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " app019CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app019DocNum").toString()) != sheet.getRow(15).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " app019DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app019SA").toString()) != sheet.getRow(15).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " app019SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("app019PG").toString()) != sheet.getRow(15).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " app019PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("tool010Past").toString()) != sheet.getRow(22).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " tool010Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool010Accept").toString()) != sheet.getRow(22).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " tool010Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool010Finish").toString()) != sheet.getRow(22).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " tool010Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool010NotFinish").toString()) != sheet.getRow(22).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " tool010NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool010CodeNum").toString()) != sheet.getRow(22).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " tool010CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool010DocNum").toString()) != sheet.getRow(22).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " tool010DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool010SA").toString()) != sheet.getRow(22).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " tool010SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool010PG").toString()) != sheet.getRow(22).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " tool010PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("tool011Past").toString()) != sheet.getRow(24).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " tool011Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool011Accept").toString()) != sheet.getRow(24).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " tool011Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool011Finish").toString()) != sheet.getRow(24).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " tool011Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool011NotFinish").toString()) != sheet.getRow(24).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " tool011NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool011CodeNum").toString()) != sheet.getRow(24).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " tool011CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool011DocNum").toString()) != sheet.getRow(24).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " tool011DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool011SA").toString()) != sheet.getRow(24).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " tool011SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool011PG").toString()) != sheet.getRow(24).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " tool011PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("tool012Past").toString()) != sheet.getRow(25).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " tool012Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool012Accept").toString()) != sheet.getRow(25).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " tool012Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool012Finish").toString()) != sheet.getRow(25).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " tool012Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool012NotFinish").toString()) != sheet.getRow(25).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " tool012NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool012CodeNum").toString()) != sheet.getRow(25).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " tool012CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool012DocNum").toString()) != sheet.getRow(25).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " tool012DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool012SA").toString()) != sheet.getRow(25).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " tool012SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool012PG").toString()) != sheet.getRow(25).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " tool012PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("tool018Past").toString()) != sheet.getRow(27).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " tool018Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool018Accept").toString()) != sheet.getRow(27).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " tool018Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool018Finish").toString()) != sheet.getRow(27).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " tool018Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool018NotFinish").toString()) != sheet.getRow(27).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " tool018NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool018CodeNum").toString()) != sheet.getRow(27).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " tool018CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool018DocNum").toString()) != sheet.getRow(27).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " tool018DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool018SA").toString()) != sheet.getRow(27).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " tool018SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool018PG").toString()) != sheet.getRow(27).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " tool018PG 數值錯誤!");

		if(Double.parseDouble(maintainRS.get("tool019Past").toString()) != sheet.getRow(28).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " tool019Past 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool019Accept").toString()) != sheet.getRow(28).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " tool019Accept 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool019Finish").toString()) != sheet.getRow(28).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " tool019Finish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool019NotFinish").toString()) != sheet.getRow(28).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " tool019NotFinish 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool019CodeNum").toString()) != sheet.getRow(28).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " tool019CodeNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool019DocNum").toString()) != sheet.getRow(28).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " tool019DocNum 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool019SA").toString()) != sheet.getRow(28).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " tool019SA 數值錯誤!");
		if(Double.parseDouble(maintainRS.get("tool019PG").toString()) != sheet.getRow(28).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " tool019PG 數值錯誤!");

		// 驗證彙總表的總計欄位內容
		if(Double.parseDouble(maintainRS.get("appCountPast").toString()) != sheet.getRow(17).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " APP 上期未完成件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("appCountAccept").toString()) != sheet.getRow(17).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " APP 受理件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("appCountFinish").toString()) != sheet.getRow(17).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " APP 完成件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("appCountNotFinish").toString()) != sheet.getRow(17).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " APP 未完成件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("appCountCodeNum").toString()) != sheet.getRow(17).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " APP 程式支數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("appCountDocNum").toString()) != sheet.getRow(17).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " APP 文件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("appCountSA").toString()) != sheet.getRow(17).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " APP 系統分析人時 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("appCountPG").toString()) != sheet.getRow(17).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " APP 程式設計人時 總計錯誤!");
		
		if(Double.parseDouble(maintainRS.get("toolCountPast").toString()) != sheet.getRow(30).getCell(2).getNumericCellValue()) throw new Exception(sheetName + " TOOL 上期未完成件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("toolCountAccept").toString()) != sheet.getRow(30).getCell(3).getNumericCellValue()) throw new Exception(sheetName + " TOOL 受理件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("toolCountFinish").toString()) != sheet.getRow(30).getCell(4).getNumericCellValue()) throw new Exception(sheetName + " TOOL 完成件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("toolCountNotFinish").toString()) != sheet.getRow(30).getCell(5).getNumericCellValue()) throw new Exception(sheetName + " TOOL 未完成件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("toolCountCodeNum").toString()) != sheet.getRow(30).getCell(10).getNumericCellValue()) throw new Exception(sheetName + " TOOL 程式支數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("toolCountDocNum").toString()) != sheet.getRow(30).getCell(11).getNumericCellValue()) throw new Exception(sheetName + " TOOL 文件數 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("toolCountSA").toString()) != sheet.getRow(30).getCell(12).getNumericCellValue()) throw new Exception(sheetName + " TOOL 系統分析人時 總計錯誤!");
		if(Double.parseDouble(maintainRS.get("toolCountPG").toString()) != sheet.getRow(30).getCell(13).getNumericCellValue()) throw new Exception(sheetName + " TOOL 程式設計人時 總計錯誤!");

	}
	
	private static Map<String, Float> chkMaintainList(String sheetName, Sheet sheet, Integer lastMonthTotalFinishNum, Integer lastMonthTotalCodeNum) throws Exception {

		Map<String,Float> map = new HashMap<String,Float>();
		Date acceptDate, replyDate, dueDate, actDate, date;
		String issueType = "", validResult = "", notFinish = "", module = "", cellValue = "", isPay = "", errMsgTitle = "";
		Float itemCount = 0f, codeNum = 0f, docNum = 0f, saNum = 0f, pgNum = 0f, totalCodeNum = 0f, totalNotFinishNum = 0f, totalFinishNum = 0f,
			app010Past = 0f, app010Accept = 0f, app010Finish = 0f, app010NotFinish = 0f, app010CodeNum = 0f, app010DocNum = 0f, app010SA = 0f, app010PG = 0f, 
			app011Past = 0f, app011Accept = 0f, app011Finish = 0f, app011NotFinish = 0f, app011CodeNum = 0f, app011DocNum = 0f, app011SA = 0f, app011PG = 0f, 
			app012Past = 0f, app012Accept = 0f, app012Finish = 0f, app012NotFinish = 0f, app012CodeNum = 0f, app012DocNum = 0f, app012SA = 0f, app012PG = 0f, 
			app013Past = 0f, app013Accept = 0f, app013Finish = 0f, app013NotFinish = 0f, app013CodeNum = 0f, app013DocNum = 0f, app013SA = 0f, app013PG = 0f, 
			app014Past = 0f, app014Accept = 0f, app014Finish = 0f, app014NotFinish = 0f, app014CodeNum = 0f, app014DocNum = 0f, app014SA = 0f, app014PG = 0f, 
			app015Past = 0f, app015Accept = 0f, app015Finish = 0f, app015NotFinish = 0f, app015CodeNum = 0f, app015DocNum = 0f, app015SA = 0f, app015PG = 0f, 
			app018Past = 0f, app018Accept = 0f, app018Finish = 0f, app018NotFinish = 0f, app018CodeNum = 0f, app018DocNum = 0f, app018SA = 0f, app018PG = 0f, 
			app019Past = 0f, app019Accept = 0f, app019Finish = 0f, app019NotFinish = 0f, app019CodeNum = 0f, app019DocNum = 0f, app019SA = 0f, app019PG = 0f, 
			tool010Past = 0f, tool010Accept = 0f, tool010Finish = 0f, tool010NotFinish = 0f, tool010CodeNum = 0f, tool010DocNum = 0f, tool010SA = 0f, tool010PG = 0f, 
			tool011Past = 0f, tool011Accept = 0f, tool011Finish = 0f, tool011NotFinish = 0f, tool011CodeNum = 0f, tool011DocNum = 0f, tool011SA = 0f, tool011PG = 0f, 
			tool012Past = 0f, tool012Accept = 0f, tool012Finish = 0f, tool012NotFinish = 0f, tool012CodeNum = 0f, tool012DocNum = 0f, tool012SA = 0f, tool012PG = 0f, 
			tool018Past = 0f, tool018Accept = 0f, tool018Finish = 0f, tool018NotFinish = 0f, tool018CodeNum = 0f, tool018DocNum = 0f, tool018SA = 0f, tool018PG = 0f, 
			tool019Past = 0f, tool019Accept = 0f, tool019Finish = 0f, tool019NotFinish = 0f, tool019CodeNum = 0f, tool019DocNum = 0f, tool019SA = 0f, tool019PG = 0f,
			appCountPast = 0f, appCountAccept = 0f, appCountFinish = 0f, appCountNotFinish = 0f, appCountCodeNum = 0f, appCountDocNum = 0f, appCountSA = 0f, appCountPG = 0f,
			toolCountPast = 0f, toolCountAccept = 0f, toolCountFinish = 0f, toolCountNotFinish = 0f, toolCountCodeNum = 0f, toolCountDocNum = 0f, toolCountSA = 0f, toolCountPG = 0f ;

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet) {
			if (row.getRowNum() >= 5 && row.getCell(0) != null) {

				acceptDate = null;
				replyDate = null;
				dueDate = null;
				actDate = null;
				date = null;
				issueType = "";
				notFinish = "";
				cellValue = "";
				isPay = "";
				codeNum = 0f;
				docNum = 0f;
				saNum = 0f;
				pgNum = 0f;
				
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
							cellValue = row.getCell(c).toString();
							notFinish += cellValue;
							if(c == 13)
								codeNum = Float.parseFloat(cellValue);
							else if(c == 14)
								docNum = Float.parseFloat(cellValue);
							else if(c == 16)
								saNum = Float.parseFloat(cellValue);
							else if(c == 17)
								pgNum = Float.parseFloat(cellValue);
						} else if (c == 18)
							isPay = row.getCell(c).toString();
					}

				}

				// 受理時間為空則表示已統計完項目資訊
				if(acceptDate == null && replyDate == null )
					break;
				
				// 計算item總數
				itemCount++;
				
				totalCodeNum += codeNum;
				
				if(actDate == null)
					totalNotFinishNum++;
				else
					totalFinishNum++;
						
				errMsgTitle = sheetName + " 序號" + (row.getRowNum() - 5 + 1) + "驗證失敗: ";
				validResult = MonthReportTools.getValidResult(acceptDate, replyDate, dueDate, actDate, issueType);
				if(validResult.startsWith("ERR"))
					throw new Exception(errMsgTitle + validResult);

				if(("app".equals(module) && !"Y".equals(isPay)) || ("tool".equals(module) && !"N".equals(isPay)))
					throw new Exception(errMsgTitle + "列入核銷結果錯誤");
				
				if ((actDate == null || (actDate != null && MonthReportTools.isFutureDate(actDate)))
						&& !"0.00.0".equals(notFinish))
					throw new Exception(errMsgTitle + "notFinish後面欄位未清空");
				
				// ============================ 計算彙總筆數 =======================================

				// 過去的單
				if(MonthReportTools.isPastDate(acceptDate)) {
					// 應用系統類別
					if ("app".equals(module)) {
						// 問題類型
						if ("010".equals(issueType)) {
							app010Past++;
							app010CodeNum += codeNum;
							app010DocNum += docNum;
							app010SA += saNum;
							app010PG += pgNum;
							if (actDate == null)
								app010NotFinish++;
							else
								app010Finish++;
						} else if ("011".equals(issueType)) {
							app011Past++;
							app011CodeNum += codeNum;
							app011DocNum += docNum;
							app011SA += saNum;
							app011PG += pgNum;
							if (actDate == null)
								app011NotFinish++;
							else
								app011Finish++;
						} else if ("012".equals(issueType)) {
							app012Past++;
							app012CodeNum += codeNum;
							app012DocNum += docNum;
							app012SA += saNum;
							app012PG += pgNum;
							if (actDate == null)
								app012NotFinish++;
							else
								app012Finish++;
						} else if ("013".equals(issueType)) {
							app013Past++;
							app013CodeNum += codeNum;
							app013DocNum += docNum;
							app013SA += saNum;
							app013PG += pgNum;
							if (actDate == null)
								app013NotFinish++;
							else
								app013Finish++;
						} else if ("014".equals(issueType)) {
							app014Past++;
							app014CodeNum += codeNum;
							app014DocNum += docNum;
							app014SA += saNum;
							app014PG += pgNum;
							if (actDate == null)
								app014NotFinish++;
							else
								app014Finish++;
						} else if ("015".equals(issueType)) {
							app015Past++;
							app015CodeNum += codeNum;
							app015DocNum += docNum;
							app015SA += saNum;
							app015PG += pgNum;
							if (actDate == null)
								app015NotFinish++;
							else
								app015Finish++;
						} else if ("018".equals(issueType)) {
							app018Past++;
							app018CodeNum += codeNum;
							app018DocNum += docNum;
							app018SA += saNum;
							app018PG += pgNum;
							if (actDate == null)
								app018NotFinish++;
							else
								app018Finish++;
						} else if ("019".equals(issueType)) {
							app019Past++;
							app019CodeNum += codeNum;
							app019DocNum += docNum;
							app019SA += saNum;
							app019PG += pgNum;
							if (actDate == null)
								app019NotFinish++;
							else
								app019Finish++;
						}
					} else { // 工具軟體類別
						if ("010".equals(issueType)) {
							tool010Past++;
							tool010CodeNum += codeNum;
							tool010DocNum += docNum;
							tool010SA += saNum;
							tool010PG += pgNum;
							if (actDate == null)
								tool010NotFinish++;
							else
								tool010Finish++;
						} else if ("011".equals(issueType)) {
							tool011Past++;
							tool011CodeNum += codeNum;
							tool011DocNum += docNum;
							tool011SA += saNum;
							tool011PG += pgNum;
							if (actDate == null)
								tool011NotFinish++;
							else
								tool011Finish++;
						} else if ("012".equals(issueType)) {
							tool012Past++;
							tool012CodeNum += codeNum;
							tool012DocNum += docNum;
							tool012SA += saNum;
							tool012PG += pgNum;
							if (actDate == null)
								tool012NotFinish++;
							else
								tool012Finish++;
						} else if ("018".equals(issueType)) {
							tool018Past++;
							tool018CodeNum += codeNum;
							tool018DocNum += docNum;
							tool018SA += saNum;
							tool018PG += pgNum;
							if (actDate == null)
								tool018NotFinish++;
							else
								tool018Finish++;
						} else if ("019".equals(issueType)) {
							tool019Past++;
							tool019CodeNum += codeNum;
							tool019DocNum += docNum;
							tool019SA += saNum;
							tool019PG += pgNum;
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
							app010CodeNum += codeNum;
							app010DocNum += docNum;
							app010SA += saNum;
							app010PG += pgNum;
							if (actDate == null)
								app010NotFinish++;
							else
								app010Finish++;
						} else if ("011".equals(issueType)) {
							app011Accept++;
							app011CodeNum += codeNum;
							app011DocNum += docNum;
							app011SA += saNum;
							app011PG += pgNum;
							if (actDate == null)
								app011NotFinish++;
							else
								app011Finish++;
						} else if ("012".equals(issueType)) {
							app012Accept++;
							app012CodeNum += codeNum;
							app012DocNum += docNum;
							app012SA += saNum;
							app012PG += pgNum;
							if (actDate == null)
								app012NotFinish++;
							else
								app012Finish++;
						} else if ("013".equals(issueType)) {
							app013Accept++;
							app013CodeNum += codeNum;
							app013DocNum += docNum;
							app013SA += saNum;
							app013PG += pgNum;
							if (actDate == null)
								app013NotFinish++;
							else
								app013Finish++;
						} else if ("014".equals(issueType)) {
							app014Accept++;
							app014CodeNum += codeNum;
							app014DocNum += docNum;
							app014SA += saNum;
							app014PG += pgNum;
							if (actDate == null)
								app014NotFinish++;
							else
								app014Finish++;
						} else if ("015".equals(issueType)) {
							app015Accept++;
							app015CodeNum += codeNum;
							app015DocNum += docNum;
							app015SA += saNum;
							app015PG += pgNum;
							if (actDate == null)
								app015NotFinish++;
							else
								app015Finish++;
						} else if ("018".equals(issueType)) {
							app018Accept++;
							app018CodeNum += codeNum;
							app018DocNum += docNum;
							app018SA += saNum;
							app018PG += pgNum;
							if (actDate == null)
								app018NotFinish++;
							else
								app018Finish++;
						} else if ("019".equals(issueType)) {
							app019Accept++;
							app019CodeNum += codeNum;
							app019DocNum += docNum;
							app019SA += saNum;
							app019PG += pgNum;
							if (actDate == null)
								app019NotFinish++;
							else
								app019Finish++;
						}
					} else { // 工具軟體類別
						if ("010".equals(issueType)) {
							tool010Accept++;
							tool010CodeNum += codeNum;
							tool010DocNum += docNum;
							tool010SA += saNum;
							tool010PG += pgNum;
							if (actDate == null)
								tool010NotFinish++;
							else
								tool010Finish++;
						} else if ("011".equals(issueType)) {
							tool011Accept++;
							tool011CodeNum += codeNum;
							tool011DocNum += docNum;
							tool011SA += saNum;
							tool011PG += pgNum;
							if (actDate == null)
								tool011NotFinish++;
							else
								tool011Finish++;
						} else if ("012".equals(issueType)) {
							tool012Accept++;
							tool012CodeNum += codeNum;
							tool012DocNum += docNum;
							tool012SA += saNum;
							tool012PG += pgNum;
							if (actDate == null)
								tool012NotFinish++;
							else
								tool012Finish++;
						} else if ("018".equals(issueType)) {
							tool018Accept++;
							tool018CodeNum += codeNum;
							tool018DocNum += docNum;
							tool018SA += saNum;
							tool018PG += pgNum;
							if (actDate == null)
								tool018NotFinish++;
							else
								tool018Finish++;
						} else if ("019".equals(issueType)) {
							tool019Accept++;
							tool019CodeNum += codeNum;
							tool019DocNum += docNum;
							tool019SA += saNum;
							tool019PG += pgNum;
							if (actDate == null)
								tool019NotFinish++;
							else
								tool019Finish++;
						}
					}
				}
				// ============================ 計算彙總筆數 END =======================================
				
			}
		}
		
		// 系統維護問題紀錄表最下方的彙總資訊
		int summaryStrRow = 5 + itemCount.intValue() + 2;
		String sheetSummaryStr1 = sheet.getRow(summaryStrRow++).getCell(0).getStringCellValue();
		String sheetSummaryStr2 = sheet.getRow(summaryStrRow).getCell(0).getStringCellValue();
		String summaryStr1 = "本月未完成件數：" + totalNotFinishNum.intValue() + " 　本月已完成件數：" + (itemCount.intValue() - totalNotFinishNum.intValue())
				+ "   　累計已完成件數：" + (lastMonthTotalFinishNum + totalFinishNum.intValue()) + "  　累計紀錄單總計："
				+ (lastMonthTotalFinishNum + itemCount.intValue());
		String summaryStr2 = "本月程式支數：" + totalCodeNum.intValue() + "  　累計程式支數：" + (lastMonthTotalCodeNum + totalCodeNum.intValue());

//System.out.println(sheetSummaryStr1);
//System.out.println(sheetSummaryStr2);
//System.out.println(summaryStr1);
//System.out.println(summaryStr2);
		if(!sheetSummaryStr1.equals(summaryStr1))
			throw new Exception(sheetName + " 彙總資訊錯誤:完成件數");
		if(!sheetSummaryStr2.equals(summaryStr2))
			throw new Exception(sheetName + " 彙總資訊錯誤:程式支數");

System.out.println(
		"\r\n, app010Past = "+app010Past+", app010Accept = "+app010Accept+", app010Finish = "+app010Finish+", app010NotFinish = "+app010NotFinish+", app010CodeNum = "+app010CodeNum+", app010DocNum = "+app010DocNum+", app010SA = "+app010SA+", app010PG = "+app010PG
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
		
		appCountPast = app010Past + app011Past + app012Past + app013Past + app014Past + app015Past + app018Past + app019Past;
		appCountAccept = app010Accept + app011Accept + app012Accept + app013Accept + app014Accept + app015Accept + app018Accept + app019Accept;
		appCountFinish = app010Finish + app011Finish + app012Finish + app013Finish + app014Finish + app015Finish + app018Finish + app019Finish;
		appCountNotFinish = app010NotFinish + app011NotFinish + app012NotFinish + app013NotFinish + app014NotFinish + app015NotFinish + app018NotFinish + app019NotFinish;
		appCountCodeNum = app010CodeNum + app011CodeNum + app012CodeNum + app013CodeNum + app014CodeNum + app015CodeNum + app018CodeNum + app019CodeNum;
		appCountDocNum = app010DocNum + app011DocNum + app012DocNum + app013DocNum + app014DocNum + app015DocNum + app018DocNum + app019DocNum;
		appCountSA = app010SA + app011SA + app012SA + app013SA + app014SA + app015SA + app018SA + app019SA;
		appCountPG = app010PG + app011PG + app012PG + app013PG + app014PG + app015PG + app018PG + app019PG;
		
		toolCountPast = tool010Past + tool011Past + tool012Past +  tool018Past + tool019Past;
		toolCountAccept = tool010Accept + tool011Accept + tool012Accept + tool018Accept + tool019Accept;
		toolCountFinish = tool010Finish + tool011Finish + tool012Finish + tool018Finish + tool019Finish;
		toolCountNotFinish = tool010NotFinish + tool011NotFinish + tool012NotFinish + tool018NotFinish + tool019NotFinish;
		toolCountCodeNum = tool010CodeNum + tool011CodeNum + tool012CodeNum + tool018CodeNum + tool019CodeNum;
		toolCountDocNum = tool010DocNum + tool011DocNum + tool012DocNum + tool018DocNum + tool019DocNum;
		toolCountSA = tool010SA + tool011SA + tool012SA + tool018SA + tool019SA;
		toolCountPG = tool010PG + tool011PG + tool012PG + tool018PG + tool019PG;
		

		map.put("appCountPast", appCountPast);
		map.put("appCountAccept", appCountAccept);
		map.put("appCountFinish", appCountFinish);
		map.put("appCountNotFinish", appCountNotFinish);
		map.put("appCountCodeNum", appCountCodeNum);
		map.put("appCountDocNum", appCountDocNum);
		map.put("appCountSA", appCountSA);
		map.put("appCountPG", appCountPG);		

		map.put("toolCountPast", toolCountPast);
		map.put("toolCountAccept", toolCountAccept);
		map.put("toolCountFinish", toolCountFinish);
		map.put("toolCountNotFinish", toolCountNotFinish);
		map.put("toolCountCodeNum", toolCountCodeNum);
		map.put("toolCountDocNum", toolCountDocNum);
		map.put("toolCountSA", toolCountSA);
		map.put("toolCountPG", toolCountPG);
		
		map.put("app010Past", app010Past);
		map.put("app010Accept", app010Accept);
		map.put("app010Finish", app010Finish);
		map.put("app010NotFinish", app010NotFinish);
		map.put("app010CodeNum", app010CodeNum);
		map.put("app010DocNum", app010DocNum);
		map.put("app010SA", app010SA);
		map.put("app010PG", app010PG);
		map.put("app011Past", app011Past);
		map.put("app011Accept", app011Accept);
		map.put("app011Finish", app011Finish);
		map.put("app011NotFinish", app011NotFinish);
		map.put("app011CodeNum", app011CodeNum);
		map.put("app011DocNum", app011DocNum);
		map.put("app011SA", app011SA);
		map.put("app011PG", app011PG);
		map.put("app012Past", app012Past);
		map.put("app012Accept", app012Accept);
		map.put("app012Finish", app012Finish);
		map.put("app012NotFinish", app012NotFinish);
		map.put("app012CodeNum", app012CodeNum);
		map.put("app012DocNum", app012DocNum);
		map.put("app012SA", app012SA);
		map.put("app012PG", app012PG);
		map.put("app013Past", app013Past);
		map.put("app013Accept", app013Accept);
		map.put("app013Finish", app013Finish);
		map.put("app013NotFinish", app013NotFinish);
		map.put("app013CodeNum", app013CodeNum);
		map.put("app013DocNum", app013DocNum);
		map.put("app013SA", app013SA);
		map.put("app013PG", app013PG);
		map.put("app014Past", app014Past);
		map.put("app014Accept", app014Accept);
		map.put("app014Finish", app014Finish);
		map.put("app014NotFinish", app014NotFinish);
		map.put("app014CodeNum", app014CodeNum);
		map.put("app014DocNum", app014DocNum);
		map.put("app014SA", app014SA);
		map.put("app014PG", app014PG);
		map.put("app015Past", app015Past);
		map.put("app015Accept", app015Accept);
		map.put("app015Finish", app015Finish);
		map.put("app015NotFinish", app015NotFinish);
		map.put("app015CodeNum", app015CodeNum);
		map.put("app015DocNum", app015DocNum);
		map.put("app015SA", app015SA);
		map.put("app015PG", app015PG);
		map.put("app018Past", app018Past);
		map.put("app018Accept", app018Accept);
		map.put("app018Finish", app018Finish);
		map.put("app018NotFinish", app018NotFinish);
		map.put("app018CodeNum", app018CodeNum);
		map.put("app018DocNum", app018DocNum);
		map.put("app018SA", app018SA);
		map.put("app018PG", app018PG);
		map.put("app019Past", app019Past);
		map.put("app019Accept", app019Accept);
		map.put("app019Finish", app019Finish);
		map.put("app019NotFinish", app019NotFinish);
		map.put("app019CodeNum", app019CodeNum);
		map.put("app019DocNum", app019DocNum);
		map.put("app019SA", app019SA);
		map.put("app019PG", app019PG);
		map.put("tool010Past", tool010Past);
		map.put("tool010Accept", tool010Accept);
		map.put("tool010Finish", tool010Finish);
		map.put("tool010NotFinish", tool010NotFinish);
		map.put("tool010CodeNum", tool010CodeNum);
		map.put("tool010DocNum", tool010DocNum);
		map.put("tool010SA", tool010SA);
		map.put("tool010PG", tool010PG);
		map.put("tool011Past", tool011Past);
		map.put("tool011Accept", tool011Accept);
		map.put("tool011Finish", tool011Finish);
		map.put("tool011NotFinish", tool011NotFinish);
		map.put("tool011CodeNum", tool011CodeNum);
		map.put("tool011DocNum", tool011DocNum);
		map.put("tool011SA", tool011SA);
		map.put("tool011PG", tool011PG);
		map.put("tool012Past", tool012Past);
		map.put("tool012Accept", tool012Accept);
		map.put("tool012Finish", tool012Finish);
		map.put("tool012NotFinish", tool012NotFinish);
		map.put("tool012CodeNum", tool012CodeNum);
		map.put("tool012DocNum", tool012DocNum);
		map.put("tool012SA", tool012SA);
		map.put("tool012PG", tool012PG);
		map.put("tool018Past", tool018Past);
		map.put("tool018Accept", tool018Accept);
		map.put("tool018Finish", tool018Finish);
		map.put("tool018NotFinish", tool018NotFinish);
		map.put("tool018CodeNum", tool018CodeNum);
		map.put("tool018DocNum", tool018DocNum);
		map.put("tool018SA", tool018SA);
		map.put("tool018PG", tool018PG);
		map.put("tool019Past", tool019Past);
		map.put("tool019Accept", tool019Accept);
		map.put("tool019Finish", tool019Finish);
		map.put("tool019NotFinish", tool019NotFinish);
		map.put("tool019CodeNum", tool019CodeNum);
		map.put("tool019DocNum", tool019DocNum);
		map.put("tool019SA", tool019SA);
		map.put("tool019PG", tool019PG);
		map.put("itemCount", itemCount);
		
		return map;

	}

}
