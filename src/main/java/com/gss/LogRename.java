package com.gss;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class LogRename {

	public static void logRename(String path, List<Map<String, String>> qcLogList, String downloadsPath) {
		File f;
		FileOutputStream fos = null;
		Cell cell;
		Workbook workbook = null;
		Sheet sheet;
		String run_flag = "", rqId = "", excelPath = "", filePath = "", txt = "" ;

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
						
						run_flag = cell.getCellType() == Cell.CELL_TYPE_NUMERIC
								? String.valueOf(row.getCell(10).getNumericCellValue())
								: row.getCell(10).getStringCellValue();

						// 將失敗的RQ Log加上run_flag結果
						f = new File(filePath + rqId + ".log");
						if (f.exists())
							f.renameTo(new File(filePath + run_flag + "_" + rqId + ".log"));
						
						// 將失敗的RQ另外整理寫入jobF.txt
						if(run_flag.contains("3"))
							txt += map.get("qcLogFile") + "\t" + run_flag + "\t" + rqId + ".log \r\n";
						
					}
				}
				workbook.close();
				System.out.println(filePath + "logRename Done");
			}

			// 將失敗的RQ資料寫入file中
			Tools.writeListFtoFile(path, txt);
			
			System.out.println("All logRename Done");
		} catch (Exception ex) {
			System.out.println("logRename Error: ");
			ex.printStackTrace();
		} finally {
			try {
				if (workbook != null)
					workbook.close();
				if (fos != null)
					fos.close();
			} catch (IOException e) {
				System.out.println("logRename finally Error: ");
				e.printStackTrace();
			}
		}
	}

}
