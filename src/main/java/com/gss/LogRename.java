package com.gss;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class LogRename {

	public static void logRename(List<Map<String, String>> qcLogList, String downloadsPath) {
		File f;
		Cell cell;
		Workbook workbook = null;
		Sheet sheet;
		String run_flag = "", rqId = "", excelPath = "", filePath = "";

		try {
			for (Map<String, String> map : qcLogList) {
				filePath = downloadsPath + map.get("qcLogFile") + "/";
				excelPath = filePath + map.get("qcLogExcel");

				f = new File(excelPath);
				workbook = Tools.getWorkbook(excelPath, f);
				sheet = workbook.getSheetAt(2);

				for (Row row : sheet) {
					if (row.getRowNum() > 1 && row.getCell(0) != null) {
						rqId = row.getCell(2).getStringCellValue();
						cell = row.getCell(10);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
							run_flag = String.valueOf(row.getCell(10).getNumericCellValue());
						else
							run_flag = row.getCell(10).getStringCellValue();
						f = new File(filePath + rqId + ".log");
						if (f.exists())
							f.renameTo(new File(filePath + run_flag + "_" + rqId + ".log"));

					}
				}
				workbook.close();
				System.out.println(filePath + "logRename Done");
			}
			System.out.println("All logRename Done");
		} catch (Exception ex) {
			System.out.println("logRename Error: ");
			ex.printStackTrace();
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				System.out.println("logRename finally Error: ");
				e.printStackTrace();
			}
		}
	}

}
