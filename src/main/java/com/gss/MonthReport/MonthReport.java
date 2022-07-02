package com.gss.MonthReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.gss.DailyReportMain;
import com.gss.Property;
import com.gss.Tools;

public class MonthReport {

//	public static void main(String[] args) {
//		String os = System.getProperty("os.name");
//		System.out.println("=== NOW TIME ===> " + new Date());
//		System.out.println("===os.name===> " + os);
//
//		// 判斷當前執行的啟動方式是IDE還是jar
//		boolean isStartupFromJar = new File(
//				DailyReportMain.class.getProtectionDomain().getCodeSource().getLocation().getPath()).isFile();
//		System.out.println("isStartupFromJar: " + isStartupFromJar);
//
//		String path = System.getProperty("user.dir") + File.separator; // Jar
//		if (!isStartupFromJar) // IDE
//			path = os.contains("Mac") ? "/Users/nicole/Dropbox/DailyReport/" // Mac
//					: "C:/Users/Nicole/Dropbox/DailyReport/"; // win
//		monthReport(path);
//	}

	static SimpleDateFormat sdfDateTime = new SimpleDateFormat("yyyy/MM/dd HH:mm");
	static SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy/MM/dd");
	static SimpleDateFormat sdfMM = new SimpleDateFormat("yyyyMM");
	static DecimalFormat df = new DecimalFormat("#.#");

	public static void monthReport(String path) {
		XSSFWorkbook xssfWorkbook = null;
		OutputStream output = null;

		System.out.println("path: " + path);
		Map<String, String> mapProp = Property.getProperties(path);

		// 月報放置路徑與檔名
		String MonthReportExcel = mapProp.get("MonthReportExcel"); // Jar
		String DailyReportExcel = path + MonthReportExcel; // Jar
		System.out.println("月報Excel: " + DailyReportExcel);
		try {
			File f = new File(DailyReportExcel);
			InputStream inputStream = new FileInputStream(f);
			// XSSF (.xlsx)
			xssfWorkbook = new XSSFWorkbook(inputStream);
			/**
			 * 因remove完後getNumberOfSheets的結果就會-1，原本的sheet3也會變成sheet2，故從尾開始跑
			 */
			for (int i = xssfWorkbook.getNumberOfSheets(); i > 0; i--) {
				if (i > 2)
					xssfWorkbook.removeSheetAt(i - 1);
			}
			setJira(xssfWorkbook);
			setGoogleDoc(xssfWorkbook);
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

	private static void setJira(XSSFWorkbook xssfWorkbook) throws Exception {

		List<List<String>> listRow = new LinkedList<List<String>>();
		List<String> listCell;
		XSSFSheet sheet1 = xssfWorkbook.getSheetAt(0);

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet1) {
			if (row.getRowNum() > 3 && row.getCell(0) != null && row.getRowNum() < sheet1.getLastRowNum()) {
				listCell = new LinkedList<String>();
				for (int c = 0; c < row.getLastCellNum(); c++) {
					if (row.getCell(c) != null && row.getCell(c).toString().length() > 0) {
						if (c == 0 || c == 1)
							listCell.add(sdfDateTime.format(row.getCell(c).getDateCellValue()));
						else if (c == 4)
							listCell.add(row.getCell(c).getStringCellValue().substring(0, 3));
						else if (c == 6 || c == 7)
							listCell.add(sdfDate.format(row.getCell(c).getDateCellValue()));
						else
							listCell.add(row.getCell(c).toString());
					} else
						listCell.add("");
				}
				listRow.add(listCell);
			}
		}

		XSSFSheet sheet = xssfWorkbook.createSheet("Jira " + sdfMM.format(new Date()));
		XSSFCellStyle style = setStyle(xssfWorkbook);
		XSSFDataFormat xssfdf = xssfWorkbook.createDataFormat();
		short dfDecimal = xssfdf.getFormat("#,#0.0"); // 貨幣
		short dfNormal = xssfdf.getFormat(""); // 通用
		short dfDate = xssfdf.getFormat("yyyy/MM/dd");
		short dfDateTime = xssfdf.getFormat("yyyy/MM/dd hh:mm");

		sheet.setDefaultColumnWidth(5);
		Row row;
		int rownum = 0, colnum = 0, dataValuenum = 0;
		String sa = "", pg = "";

		for (List<String> dataCell : listRow) {
			row = sheet.createRow(rownum++);
			colnum = 0;
			dataValuenum = 0;

			for (String dataValue : dataCell) {

				dataValuenum++;
				if (dataValuenum == 1) { // 序號
					addCellStyle(row, colnum++, style, dfNormal).setCellFormula("ROW()-5");
				} else if (dataValuenum == 3) { // 逾期 * 2
					for (int i = 0; i < 2; i++) {
						addCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
					}
				} else if (dataValuenum == 9) { // 逾期 * 2 , 數量 * 2
					for (int i = 0; i < 4; i++) {
						addCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
					}
				} else if (dataValuenum == 10) { // 人時 * 2, 核銷
					sa = "0";
					pg = "0";
					if (dataValue.length() > 0) {
						dataValue = dataValue.replaceAll("[ -]", "").toUpperCase();
						if (dataValue.indexOf("SA") >= 0)
							sa = dataValue.substring(dataValue.indexOf("SA") + 2, dataValue.indexOf("H"));
						if (dataValue.indexOf("PG") >= 0)
							pg = dataValue.substring(dataValue.lastIndexOf("PG") + 2, dataValue.lastIndexOf("H"));
					}
					addCellStyle(row, colnum++, style, dfDecimal).setCellValue(df.parse(sa).doubleValue());
					addCellStyle(row, colnum++, style, dfDecimal).setCellValue(df.parse(pg).doubleValue());
					addCellStyle(row, colnum++, style, dfNormal).setCellValue("Y");
				}

				sheet.setColumnWidth(colnum, 20 * 256);
				if (dataValuenum == 1 || dataValuenum == 2)
					addCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataValue);
				else if (dataValuenum == 7 || dataValuenum == 8)
					addCellStyle(row, colnum++, style, dfDate).setCellValue(dataValue);
				else if (dataValuenum < 10)
					addCellStyle(row, colnum++, style, dfNormal).setCellValue(dataValue);
			}
		}
	}

	private static void setGoogleDoc(XSSFWorkbook xssfWorkbook) throws Exception {

		List<List<String>> listRow = new LinkedList<List<String>>();
		List<String> listCell;
		XSSFSheet sheet1 = xssfWorkbook.getSheetAt(1);

		// 取得對應Job的Row位置(橫列)
		for (Row row : sheet1) {
			if (row.getRowNum() > 0 && row.getCell(0) != null) {
				listCell = new LinkedList<String>();
				for (int c = 0; c < row.getLastCellNum(); c++) {
					if (row.getCell(c) != null && row.getCell(c).toString().length() > 0) {
						if (c == 4 || c == 5)
							listCell.add(sdfDateTime.format(row.getCell(c).getDateCellValue()));
						else if (c == 12 || c == 13)
							listCell.add(String.valueOf(((Double) row.getCell(c).getNumericCellValue()).intValue()));
						else if (c == 15)
							listCell.add(Tools.setLen(
									String.valueOf(((Double) row.getCell(c).getNumericCellValue()).intValue()), 3));
						else if (c == 6 || c == 7)
							listCell.add(sdfDate.format(row.getCell(c).getDateCellValue()));
						else
							listCell.add(row.getCell(c).toString());
					} else
						listCell.add("");
				}
				listRow.add(listCell);
			}
		}

		XSSFSheet sheet = xssfWorkbook.createSheet("Google doc " + sdfMM.format(new Date()));
		XSSFCellStyle style = setStyle(xssfWorkbook);
		XSSFDataFormat xssfdf = xssfWorkbook.createDataFormat();
		short dfDecimal = xssfdf.getFormat("#,#0.0"); // 貨幣
		short dfNormal = xssfdf.getFormat(""); // 通用
		short dfDate = xssfdf.getFormat("yyyy/MM/dd");
		short dfDateTime = xssfdf.getFormat("yyyy/MM/dd hh:mm");

		sheet.setDefaultColumnWidth(5);
		Row row;
		int rownum = 0, colnum = 0;

		for (List<String> dataCell : listRow) {
			row = sheet.createRow(rownum++);
			colnum = 0;

			addCellStyle(row, colnum++, style, dfNormal).setCellFormula("ROW()-5");
			sheet.setColumnWidth(colnum, 20 * 256);
			addCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(4));
			sheet.setColumnWidth(colnum, 20 * 256);
			addCellStyle(row, colnum++, style, dfDateTime).setCellValue(dataCell.get(5));
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(14));
			addCellStyle(row, colnum++, style, dfNormal).setCellValue("-");
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(Tools.setLen(dataCell.get(15), 3));
			sheet.setColumnWidth(colnum, 20 * 256);
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(2));
			sheet.setColumnWidth(colnum, 20 * 256);
			addCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(6));
			sheet.setColumnWidth(colnum, 20 * 256);
			addCellStyle(row, colnum++, style, dfDate).setCellValue(dataCell.get(7));
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(0);
			addCellStyle(row, colnum++, style, dfNormal)
					.setCellValue(Integer.parseInt(dataCell.get(12).toString().length() > 0 ? dataCell.get(12) : "0"));
			addCellStyle(row, colnum++, style, dfNormal)
					.setCellValue(Integer.parseInt(dataCell.get(13).toString().length() > 0 ? dataCell.get(13) : "0"));
			sheet.setColumnWidth(colnum, 20 * 256);
			addCellStyle(row, colnum++, style, dfNormal).setCellValue(dataCell.get(11));
			addCellStyle(row, colnum++, style, dfDecimal)
					.setCellValue(Float.parseFloat(dataCell.get(9).toString().length() > 0 ? dataCell.get(9) : "0"));
			addCellStyle(row, colnum++, style, dfDecimal)
					.setCellValue(Float.parseFloat(dataCell.get(10).toString().length() > 0 ? dataCell.get(10) : "0"));
			addCellStyle(row, colnum++, style, dfNormal).setCellValue("Y");

		}
	}

	private static XSSFCellStyle setStyle(XSSFWorkbook xssfWorkbook) {
		XSSFFont font = xssfWorkbook.createFont();
		font.setFontHeightInPoints((short) 10);
		font.setFontName("微軟正黑體");

		XSSFCellStyle style = xssfWorkbook.createCellStyle();
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP); // 垂直置上
		style.setAlignment(XSSFCellStyle.ALIGN_LEFT); // 水平置左
		style.setBorderTop(XSSFCellStyle.BORDER_THIN); // 上邊框
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下邊框
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN); // 左邊框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN); // 右邊框
		style.setWrapText(true); // 自動換行
		style.setFont(font);

		return style;
	}

	private static Cell addCellStyle(Row row, int c, XSSFCellStyle style, short xssfdf) {
		Cell cell = row.createCell(c);
		XSSFCellStyle newStyle = style;
		if (xssfdf > 1) {
			newStyle = (XSSFCellStyle) style.clone();
			newStyle.setDataFormat(xssfdf);
		}
		cell.setCellStyle(newStyle);
		return cell;
	}
}
