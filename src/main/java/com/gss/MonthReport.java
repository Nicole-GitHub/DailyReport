package com.gss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MonthReport {

	public static void main(String[] args) {
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
		monthReport(path);
	}
	protected static void monthReport(String path) {
		XSSFWorkbook xssfWorkbook = null;
		OutputStream output = null;

		System.out.println("path: " + path);
		Map<String, String> mapProp = Property.getProperties(path);
		List<List<String>> listRow = new LinkedList<List<String>>();
		List listCell ;
		
		// 月報放置路徑與檔名
		String MonthReportExcel = mapProp.get("MonthReportExcel"); // Jar
		String DailyReportExcel = path + MonthReportExcel; // Jar
		System.out.println("月報Excel: " + DailyReportExcel);
		SimpleDateFormat sdfDateTime = new SimpleDateFormat("yyyy/MM/dd HH:mm");
		SimpleDateFormat sdfDate = new SimpleDateFormat("yyyy/MM/dd");
		List<Integer> chageCellFormate = Arrays.asList(0,1,6,7);
		
		try {
			File f = new File(DailyReportExcel);
			InputStream inputStream = new FileInputStream(f);
			// XSSF (.xlsx)
			xssfWorkbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet1 = xssfWorkbook.getSheetAt(0);
			
			// 取得對應Job的Row位置(橫列)
			for (Row row : sheet1) {
				if (row.getRowNum() > 3 && row.getCell(0) != null && row.getRowNum() < sheet1.getLastRowNum()) {
					listCell = new LinkedList<String>();
					for (int c = 0 ; c < row.getLastCellNum() ; c++) {
						if(row.getCell(c) != null && row.getCell(c).toString().length() > 0) {
							if (c == 0 || c == 1)
								listCell.add("'" + sdfDateTime.format(row.getCell(c).getDateCellValue()));
							else if(c == 4)
								listCell.add("'" + row.getCell(c).toString().substring(0,3));
							else if (c == 6 || c == 7)
								listCell.add("'" + sdfDate.format(row.getCell(c).getDateCellValue()));
							else
								listCell.add(row.getCell(c).toString());
						}
						else
							listCell.add("");
					}
					listRow.add(listCell);
				}
			}

			XSSFSheet sheet3 = xssfWorkbook.createSheet("jira done");
			Row sheetRow;
			int r = 0, c = 0;
			for (List<String> row : listRow) {
				sheetRow = sheet3.createRow(r++);
				c = 0;
				for (String cell : row) {
					sheetRow.createCell(c++).setCellValue(cell);
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
