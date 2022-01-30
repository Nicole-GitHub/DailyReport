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
		Workbook workbook = null;
		Sheet sheet;
		String run_flag = "", rq_id = "", excelPath = "", filePath = "", txt = "",
				control_id = "", source_tablenm = "", target_tablenm = "", exec_sdate = "", exec_edate = "";

		try {
			
			for (Map<String, String> map : qcLogList) {
				filePath = downloadsPath + map.get("qcLogFile") + "/";
				excelPath = filePath + map.get("qcLogExcel");

				f = new File(excelPath);
				workbook = Tools.getWorkbook(excelPath, f);
				sheet = workbook.getSheetAt(2);

				for (Row row : sheet) {
					if (row.getRowNum() > 1 && row.getCell(0) != null) {
						control_id = row.getCell(1).getCellType() == Cell.CELL_TYPE_NUMERIC
								? String.valueOf(row.getCell(1).getNumericCellValue())
								: row.getCell(1).getStringCellValue();
						rq_id = row.getCell(2).getStringCellValue();
						source_tablenm = row.getCell(7).getStringCellValue();
						target_tablenm = row.getCell(8).getStringCellValue();
						run_flag = row.getCell(10).getCellType() == Cell.CELL_TYPE_NUMERIC
								? String.valueOf(row.getCell(10).getNumericCellValue())
								: row.getCell(10).getStringCellValue();
						exec_sdate = row.getCell(18).getStringCellValue();
						exec_edate = row.getCell(19).getStringCellValue();

						// 將失敗的RQ Log加上run_flag結果
						f = new File(filePath + rq_id + ".log");
						if (f.exists())
							f.renameTo(new File(filePath + run_flag + "_" + rq_id + ".log"));

						// 將失敗的RQ另外整理寫入jobF.txt
						if (run_flag.contains("3"))
							txt += "JOB:\t" + map.get("qcLogFile") + " \r\n"
									+ "RQ_ID:\t" + control_id + "\t" + rq_id + "\t" + run_flag + " \r\n"
									+ "Table:\t" + source_tablenm + " -> " + target_tablenm + " \r\n"
									+ "DateTime:\t" + exec_sdate + " -> " + exec_edate + " \r\n\r\n";

					}
				}
				workbook.close();
				System.out.println(filePath + "logRename Done");
			}

			// 將失敗的RQ資料寫入file中，並做結尾
			Tools.writeListFtoFile(path, "失敗的RQ \r\n\r\n" + txt, true);
			
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
