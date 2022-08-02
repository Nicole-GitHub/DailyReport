package com.gss.MonthReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.gss.Property;
import com.gss.Tools;

public class MaintainList {

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
		String maintainListExcel = path + monthReportPath + mapProp.get("MaintainListExcel"); // Jar
		String monthReportExcel = path + monthReportPath + mapProp.get("MonthReportExcel"); // Jar
		System.out.println("維護問題紀錄單Excel: " + maintainListExcel);
		System.out.println("月報Excel: " + monthReportExcel);

		System.out.println(new Date());
		maintainList(maintainListExcel);
		System.out.println(new Date());
	}

	private static final SimpleDateFormat sdfDateTime = new SimpleDateFormat("yyyy/MM/dd HH:mm");
	private static final SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy/MM/dd");
	private static final SimpleDateFormat sdfMM = new SimpleDateFormat("yyyyMM");
	private static final DecimalFormat df = new DecimalFormat("#.#");
	private static Date acceptDate, replyDate, dueDate, actDate, date;
	private static String issueType, validResult, styleType;
	private static String[] validResultArr;
	private static long diffrence;
	private static final TimeUnit time = TimeUnit.MINUTES;
	private static final List<String> listIssueType1 = Arrays.asList(new String[] { "010" });
	private static final List<String> listIssueType2 = Arrays.asList(new String[] { "011", "012" });
	private static final List<String> listIssueType3 = Arrays
			.asList(new String[] { "013", "014", "015", "018", "019" });
	private static final List<String> listToolsModule = Arrays
			.asList(new String[] { "SAFENET", "DATASTAGE", "HADOOP", "WEBFOCUS", "IA" });
	private static List<String> listCell;
	private static XSSFCellStyle style, errStyle, normalStyle, titleStyle;
	private static short dfDecimal, dfNormal, dfDate, dfDateTime;
	private static Row row;
	

	private static void setDataFormat(XSSFWorkbook xssfWorkbook) {
		XSSFDataFormat xssfdf = xssfWorkbook.createDataFormat();
		dfDecimal = xssfdf.getFormat("#,#0.0"); // 貨幣
		dfNormal = xssfdf.getFormat(""); // 通用
		dfDate = xssfdf.getFormat("yyyy/MM/dd");
		dfDateTime = xssfdf.getFormat("yyyy/MM/dd hh:mm");
	}
	
	private static void setStyleInit(XSSFWorkbook xssfWorkbook) {
		errStyle = setStyle(xssfWorkbook, "ERR");
		normalStyle = setStyle(xssfWorkbook, "Normal");
		titleStyle = setStyle(xssfWorkbook, "Title");
	}


	private static void setVarInit() {

		listCell = new LinkedList<String>();
		acceptDate = null;
		replyDate = null;
		dueDate = null;
		actDate = null;
		date = null;
		issueType = "";
	}

	public static void maintainList(String maintainListExcel) {
		XSSFWorkbook xssfWorkbook = null;
		OutputStream output = null;
		int itemCount = 0;
		
		try {
			File f = new File(maintainListExcel);
			InputStream inputStream = new FileInputStream(f);
			// XSSF (.xlsx)
			xssfWorkbook = new XSSFWorkbook(inputStream);
			setDataFormat(xssfWorkbook);
			setStyleInit(xssfWorkbook);
			
			itemCount = setJira(xssfWorkbook,itemCount);
System.out.println("Jira Done !");

			itemCount = setGoogleDoc(xssfWorkbook,itemCount);
System.out.println("GoogleDoc Done !");

			itemCount = setIA(xssfWorkbook, 2,itemCount);
System.out.println("IA Done !");
			
			if(xssfWorkbook.getNumberOfSheets() > 3) {
				itemCount = setIA(xssfWorkbook, 3,itemCount);
System.out.println("IA2 Done !");
			}

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
		
		System.out.println("Total Item : " + itemCount);
	}

	private static int setJira(XSSFWorkbook xssfWorkbook, int itemCount) throws Exception {

		List<List<String>> listRow = new LinkedList<List<String>>();
		XSSFSheet sheet = xssfWorkbook.getSheetAt(0);

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet) {
			if (row.getRowNum() > 3 && row.getCell(0) != null && row.getRowNum() < sheet.getLastRowNum()) {
				itemCount++;
				setVarInit();

				for (int c = 0; c < row.getLastCellNum(); c++) {
					if (row.getCell(c) != null && row.getCell(c).toString().length() > 0) {
						if (c == 0 || c == 1) {
							date = row.getCell(c).getDateCellValue();
							listCell.add(sdfDateTime.format(date));
							if (c == 0)
								acceptDate = date;
							if (c == 1)
								replyDate = date;
						} else if (c == 4) {
							issueType = row.getCell(c).getStringCellValue().substring(0, 3);
							listCell.add(issueType);
						} else if (c == 6 || c == 7) {
							date = row.getCell(c).getDateCellValue();
							listCell.add(sdfDate.format(date));
							if (c == 6)
								dueDate = date;
							if (c == 7)
								actDate = date;
						} else
							listCell.add(row.getCell(c).toString());
					} else
						listCell.add("");
				}

				listCell.add(getValidResult(acceptDate, replyDate, dueDate, actDate,
						listCell.get(listCell.size() - 1).toUpperCase()));
				listRow.add(listCell);
			}
		}

		sheet = xssfWorkbook.createSheet("Jira " + sdfMM.format(new Date()));
		int rownum = 0, colnum = 0;

		row = sheet.createRow(rownum++);
		setTitleRow(xssfWorkbook, row, sheet);
		for (List<String> dataCell : listRow) {
			row = sheet.createRow(rownum++);
			colnum = 0;

			validResult = dataCell.get(dataCell.size() - 1);
			validResultArr = validResult.split(",");
			styleType = validResultArr[0].indexOf("-") >= 0
					? validResultArr[0].substring(0, validResultArr[0].indexOf("-"))
					: validResultArr[0];
			style = "ERR".equals(styleType) ? errStyle : normalStyle;

			setCellStyle(row, colnum++, style, dfNormal).setCellFormula("ROW()-5");
			setCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(0));
			setCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(1));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(2));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(3));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(4));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(5));
			setCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(6));
			setCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(7));

			if (validResultArr.length > 1) {
				for (int i = 0; i < 5; i++)
					setCellStyle(row, colnum++, style, dfNormal).setCellValue("");

				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(0.0f);
				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(0.0f);
			} else {
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(8));

				String dataValue = dataCell.get(9);
				String sa = "0", pg = "0";
				if (dataValue.length() > 0) {
					dataValue = dataValue.replaceAll("[ -]", "").toUpperCase();
					if (dataValue.indexOf("SA") >= 0)
						sa = dataValue.substring(dataValue.indexOf("SA") + 2, dataValue.indexOf("H"));
					if (dataValue.indexOf("PG") >= 0)
						pg = dataValue.substring(dataValue.lastIndexOf("PG") + 2, dataValue.lastIndexOf("H"));
				}

				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(df.parse(sa).doubleValue());
				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(df.parse(pg).doubleValue());
			}

			setCellStyle(row, colnum++, style, dfNormal)
					.setCellValue(listToolsModule.contains(dataCell.get(2).toUpperCase()) ? "N" : "Y");
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(validResult);
		}
		
		return itemCount;
	}

	private static int setGoogleDoc(XSSFWorkbook xssfWorkbook, int itemCount) throws Exception {

		List<List<String>> listRow = new LinkedList<List<String>>();
		XSSFSheet sheet = xssfWorkbook.getSheetAt(1);

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet) {
			if (row.getRowNum() > 0 && row.getCell(0) != null) {
				itemCount++;

				setVarInit();

				for (int c = 0; c < row.getLastCellNum(); c++) {
					if (row.getCell(c) != null && row.getCell(c).toString().length() > 0) {
						if (c == 4 || c == 5) {
							date = row.getCell(c).getDateCellValue();
							listCell.add(sdfDateTime.format(date));
							if (c == 4)
								acceptDate = date;
							if (c == 5)
								replyDate = date;
						} else if (c == 12 || c == 13) {
							listCell.add(
									String.valueOf(((Double) row.getCell(c).getNumericCellValue()).intValue()));
						} else if (c == 15) {
							issueType = Tools.setLen(
									String.valueOf(((Double) row.getCell(c).getNumericCellValue()).intValue()), 3);
							listCell.add(issueType);
						} else if (c == 6 || c == 7) {
							date = row.getCell(c).getDateCellValue();
							listCell.add(sdfDate.format(date));
							if (c == 6)
								dueDate = date;
							if (c == 7)
								actDate = date;
						} else
							listCell.add(row.getCell(c).toString());
					} else
						listCell.add("");
				}

				listCell.add(getValidResult(acceptDate, replyDate, dueDate, actDate,
						listCell.get(listCell.size() - 1).toUpperCase()));
				listRow.add(listCell);
			}
		}

		sheet = xssfWorkbook.createSheet("Google doc " + sdfMM.format(new Date()));
		int rownum = 0, colnum = 0;

		row = sheet.createRow(rownum++);
		setTitleRow(xssfWorkbook, row, sheet);
		for (List<String> dataCell : listRow) {
			row = sheet.createRow(rownum++);
			colnum = 0;

			validResult = dataCell.get(dataCell.size() - 1);
			validResultArr = validResult.split(",");
			styleType = validResultArr[0].indexOf("-") >= 0
					? validResultArr[0].substring(0, validResultArr[0].indexOf("-"))
					: validResultArr[0];
			style = "ERR".equals(styleType) ? errStyle : normalStyle;

			setCellStyle(row, colnum++, style, dfNormal).setCellFormula("ROW()-5");
			setCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(4));
			setCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(5));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(14));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue("-");
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(15));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(2));
			setCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(6));
			setCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(7));

			if (validResultArr.length > 1) {
				for (int i = 0; i < 5; i++)
					setCellStyle(row, colnum++, style, dfNormal).setCellValue("");

				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(0.0f);
				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(0.0f);
			} else {
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
				setCellStyle(row, colnum++, style, dfNormal)
						.setCellValue(Integer.parseInt(!"".equals(dataCell.get(12)) ? dataCell.get(12) : "0"));
				setCellStyle(row, colnum++, style, dfNormal)
						.setCellValue(Integer.parseInt(!"".equals(dataCell.get(13)) ? dataCell.get(13) : "0"));
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(11));
				setCellStyle(row, colnum++, style, dfDecimal)
						.setCellValue(Float.parseFloat(!"".equals(dataCell.get(9)) ? dataCell.get(9) : "0"));
				setCellStyle(row, colnum++, style, dfDecimal)
						.setCellValue(Float.parseFloat(!"".equals(dataCell.get(10)) ? dataCell.get(10) : "0"));
			}

			setCellStyle(row, colnum++, style, dfNormal)
					.setCellValue(listToolsModule.contains(dataCell.get(14).toUpperCase()) ? "N" : "Y");
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(validResult);

		}
		
		return itemCount;
	}

	private static int setIA(XSSFWorkbook xssfWorkbook, int sheetNum, int itemCount) throws Exception {

		List<List<String>> listRow = new LinkedList<List<String>>();
		XSSFSheet sheet = xssfWorkbook.getSheetAt(sheetNum);

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet) {
			if (row.getRowNum() > 0 && row.getCell(0) != null) {
				itemCount++;
				setVarInit();
//System.out.println(row.getRowNum());
				for (int c = 0; c < row.getLastCellNum(); c++) {
					if (row.getCell(c) != null && row.getCell(c).toString().length() > 0) {
						if (c == 1 || c == 2) {
							date = row.getCell(c).getDateCellValue();
							listCell.add(sdfDateTime.format(date));
							if (c == 1)
								acceptDate = date;
							if (c == 2)
								replyDate = date;
						} else if (c == 9 || c == 10) {
							date = row.getCell(c).getDateCellValue();
							listCell.add(sdfDate.format(date));
							if (c == 9)
								dueDate = date;
							if (c == 10)
								actDate = date;
						} else
							listCell.add(row.getCell(c).toString());
					} else
						listCell.add("");

				}

				listCell.add(getValidResult(acceptDate, replyDate, dueDate, actDate,
						listCell.get(listCell.size() - 1).toUpperCase()));
				listRow.add(listCell);
			}
		}

		sheet = xssfWorkbook.createSheet("IA"+ sheetNum + " " + sdfMM.format(new Date()));
		int rownum = 0, colnum = 0;

		row = sheet.createRow(rownum++);
		setTitleRow(xssfWorkbook, row, sheet);
		for (List<String> dataCell : listRow) {
			row = sheet.createRow(rownum++);
			colnum = 0;

			validResult = dataCell.get(dataCell.size() - 1);
			validResultArr = validResult.split(",");
			styleType = validResultArr[0].indexOf("-") >= 0
					? validResultArr[0].substring(0, validResultArr[0].indexOf("-"))
					: validResultArr[0];
			style = "ERR".equals(styleType) ? errStyle : normalStyle;

			setCellStyle(row, colnum++, style, dfNormal).setCellFormula("ROW()-5");
			setCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(1));
			setCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(2));
			setCellStyle(row, colnum++, style, dfNormal)
					.setCellValue(((Double) Double.parseDouble((dataCell.get(3)))).intValue());
			setCellStyle(row, colnum++, style, dfNormal)
					.setCellValue(((Double) Double.parseDouble((dataCell.get(4)))).intValue());
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(5));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(6));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(7));
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(8));
			setCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(9));
			setCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(10));

			if (validResultArr.length > 1) {
				for (int i = 0; i < 4; i++)
					setCellStyle(row, colnum++, style, dfNormal).setCellValue("");

				setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(15));
				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(0.0f);
				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(0.0f);
			} else {
				setCellStyle(row, colnum++, style, dfNormal)
						.setCellValue(((Double) Double.parseDouble((dataCell.get(11)))).intValue());
				setCellStyle(row, colnum++, style, dfNormal)
						.setCellValue(((Double) Double.parseDouble((dataCell.get(12)))).intValue());
				setCellStyle(row, colnum++, style, dfNormal)
						.setCellValue(((Double) Double.parseDouble((dataCell.get(13)))).intValue());
				setCellStyle(row, colnum++, style, dfNormal)
						.setCellValue(((Double) Double.parseDouble((dataCell.get(14)))).intValue());
				setCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(15));
				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(Float.parseFloat(dataCell.get(16)));
				setCellStyle(row, colnum++, style, dfDecimal).setCellValue(Float.parseFloat(dataCell.get(17)));
			}

			setCellStyle(row, colnum++, style, dfNormal)
					.setCellValue(listToolsModule.contains(dataCell.get(5).toUpperCase()) ? "N" : "Y");
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(validResult);
		}
		
		return itemCount;
	}

	/**
	 * 設定Style
	 * 
	 * @param xssfWorkbook
	 * @param type
	 * @return
	 */
	private static XSSFCellStyle setStyle(XSSFWorkbook xssfWorkbook, String type) {
		
		XSSFFont font = xssfWorkbook.createFont();
		font.setFontHeightInPoints((short) ("Title".equals(type) ? 14 : 10));
		font.setFontName("微軟正黑體");
		font.setBold("Title".equals(type));

		XSSFCellStyle style = xssfWorkbook.createCellStyle();
		short borderStyle = CellStyle.BORDER_THIN;
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP); // 垂直置上
		style.setAlignment(XSSFCellStyle.ALIGN_LEFT); // 水平置左
		style.setBorderTop(borderStyle); // 上邊框
		style.setBorderBottom(borderStyle); // 下邊框
		style.setBorderLeft(borderStyle); // 左邊框
		style.setBorderRight(borderStyle); // 右邊框
		style.setWrapText(!"Title".equals(type)); // 自動換行
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND); // Cell背景色需搭配此行才會有效果
		// Cell背景色
		style.setFillForegroundColor("ERR".equals(type) ? IndexedColors.YELLOW.index
				: "Title".equals(type) ? IndexedColors.LIGHT_GREEN.index : IndexedColors.WHITE.index);
		style.setFont(font);

		return style;
	}

	/**
	 * 設定Cell Style
	 * 
	 * @param row
	 * @param c
	 * @param style
	 * @param xssfdf
	 * @return
	 */
	private static Cell setCellStyle(Row row, int c, XSSFCellStyle style, short xssfdf) {
		Cell cell = row.createCell(c);
		XSSFCellStyle newStyle = style;
		if (xssfdf > 1) {
			newStyle = (XSSFCellStyle) style.clone();
			newStyle.setDataFormat(xssfdf);
		}
		cell.setCellStyle(newStyle);
		return cell;
	}

	/**
	 * 設定title
	 * 
	 * @param xssfWorkbook
	 * @param row
	 * @param sheet
	 */
	private static void setTitleRow(XSSFWorkbook xssfWorkbook, Row row, XSSFSheet sheet) {
		int colnum = 0;

		style = titleStyle;
		List<Integer> widthCols15 = Arrays.asList(new Integer[] { 5, 9, 10 });
		List<Integer> widthCols20 = Arrays.asList(new Integer[] { 1, 2, 6, 8, 15, 19 });
		List<Integer> widthCols35 = Arrays.asList(new Integer[] { 8 });
		List<Integer> widthCols40 = Arrays.asList(new Integer[] { 15 });
		List<String> title = Arrays.asList(new String[] { "序號", "受理時間", "回應時間", "逾期", "回應逾期時數", "模組", "紀錄單號", "問題類型",
				"問題摘要", "到期日", "實際完成日", "逾期", "完成逾期天數", "程式支數", "文件數", "處理描述", "系統分析人時", "程式設計人時", "列入核銷", "驗證結果" });

		sheet.setDefaultColumnWidth(5);
		for (Integer cols : widthCols15)
			sheet.setColumnWidth(cols, 15 * 256);
		for (Integer cols : widthCols20)
			sheet.setColumnWidth(cols, 20 * 256);
		for (Integer cols : widthCols35)
			sheet.setColumnWidth(cols, 35 * 256);
		for (Integer cols : widthCols40)
			sheet.setColumnWidth(cols, 40 * 256);
		for (String str : title)
			setCellStyle(row, colnum++, style, dfNormal).setCellValue(str);
	}

	/**
	 * 驗証結果
	 * 
	 * @param acceptDate
	 * @param replyDate
	 * @param dueDate
	 * @param actDate
	 * @param manualChk
	 * @return
	 */
	@SuppressWarnings("deprecation")
	private static String getValidResult(Date acceptDate, Date replyDate, Date dueDate, Date actDate,
			String manualChk) {
		String validResult = "";
		/**
		 * 驗證內容是否有誤
		 * 受理時間 & 回應時間 & 到期日不可為空
		 * 受理時間 & 回應時間 需根據問題類型去判斷
		 * 到期日不可早於實際完成日
		 * 實際完成日為空或為未來日期則後面的資訊皆為清空，除了時數為0.0 (notFinish)
		 */
		if (replyDate == null || dueDate == null || acceptDate == null) {
			validResult = "ERR-受理時間 & 回應時間 & 到期日不可為空";
		} else {
			// 受理時間 與 回應時間 之間是否含有週休二日
			int holiday = getHoliday(acceptDate, replyDate);
			// 受理時間為中午前則當天需算1個工作天
			int isAM = acceptDate.getHours() < 12 ? 1 : 0;
			diffrence = time.convert(replyDate.getTime() - acceptDate.getTime(), TimeUnit.MILLISECONDS);
			if ((listIssueType1.contains(issueType) && diffrence / 60f / 24f > 2 + holiday - isAM)
					|| (listIssueType2.contains(issueType) && diffrence / 60f > 4)
					|| (listIssueType3.contains(issueType) && diffrence / 60f / 24f > 3 + holiday - isAM)) {
				validResult = "ERR-受理時間 & 回應時間 需根據問題類型去判斷";
			} else if (actDate != null) {
				diffrence = time.convert(actDate.getTime() - dueDate.getTime(), TimeUnit.MILLISECONDS);
				if (diffrence / 60f / 24f > 1) {
					validResult = "ERR-到期日不可早於實際完成日";
				} else
					validResult = "Normal";
			}else
				validResult = "Normal";
		}

		// 已人工確認過無誤
		if ("V".equals(manualChk))
			validResult = "manualChk";
		if (actDate == null || (actDate != null && isFutureDate(actDate)))
			validResult += ",notFinish";

		return validResult;
	}

	/**
	 * 判斷受理日與回應日之間有無假日
	 * 
	 * @param acceptDate
	 * @param replyDate
	 * @return
	 */
	private static int getHoliday(Date acceptDate, Date replyDate) {
		Calendar c = Calendar.getInstance();
		c.setTime(acceptDate);
		int acceptDateOfWeek = c.get(Calendar.DAY_OF_WEEK);
		c.setTime(replyDate);
		int replyDateOfWeek = c.get(Calendar.DAY_OF_WEEK);
		return acceptDateOfWeek > replyDateOfWeek ? 2 : 0;
	}

	/**
	 * 是否為未來日期
	 * 
	 * @param actDate
	 * @return
	 */
	private static boolean isFutureDate(Date actDate) {
		Calendar c = Calendar.getInstance();
		c.setTime(actDate);
		if (c.get(Calendar.MONTH) >= Calendar.getInstance().get(Calendar.MONTH))
			return true;
		return false;
	}
}
