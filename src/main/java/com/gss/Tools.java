package com.gss;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Tools {

	/**
	 * 取得 Excel的Workbook
	 * 
	 * @param path
	 * @return
	 */
	protected static Workbook getWorkbook(String path, File f) {
		Workbook workbook = null;
		InputStream inputStream = null;
		try {
			inputStream = new FileInputStream(f);
			String aux = path.substring(path.lastIndexOf(".") + 1);
			if ("XLS".equalsIgnoreCase(aux)) {
				workbook = new HSSFWorkbook(inputStream);
			} else if ("XLSX".equalsIgnoreCase(aux)) {
				workbook = new XSSFWorkbook(inputStream);
			} else {
				throw new Exception("檔案格式錯誤");
			}
		} catch (Exception ex) {
			// 因output時需要用到，故不可寫在finally內
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				System.out.println("Tools getWorkbook Error:");
				e.printStackTrace();
			}

			System.out.println("Tools getWorkbook Error:");
			ex.printStackTrace();
		} finally {
			try {
				if (inputStream != null)
					inputStream.close();
			} catch (IOException e) {
				System.out.println("Tools getWorkbook Error:");
				e.printStackTrace();
			}
		}
		return workbook;
	}

	/**
	 * 取得對應日期的Cell位置(縱列)
	 * 
	 * @return
	 */
	protected static Integer getDateCell(Sheet sheet1, String JobDate) {
		for (Cell cell : sheet1.getRow(0)) {
			if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
				if (cell.getNumericCellValue() == Double.valueOf(JobDate))
					return cell.getColumnIndex();
			} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
				if (cell.getStringCellValue().equals(JobDate))
					return cell.getColumnIndex();
			}
		}
		return 0;
	}

	protected static void setCellStyle(int setColNum, Cell cell, CellStyle cellStyle, Row row, Sheet sheet3,
			Sheet sheet4, String desc) {
		cell = row.createCell(setColNum);
		cell.setCellValue(desc);
		cellStyle.cloneStyleFrom(sheet4.getRow(1).getCell(setColNum).getCellStyle());
		cell.setCellStyle(cellStyle);
	}

	/**
	 * 取得檢查的天數(今日與檢查起日的相差天數，平日會是同一天，只有假日才會有差別)
	 * 應包含今日與檢查起日，故相減後需+1
	 * 
	 * @throws ParseException
	 */
	protected static int getMinusDays(int chkDate) throws ParseException {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");

		Calendar before = Calendar.getInstance();// 檢查日
		Calendar after = Calendar.getInstance();// 今日
		before.setTime(sdf.parse(String.valueOf(chkDate)));
		int minusDays = after.get(Calendar.DATE) - before.get(Calendar.DATE);
//		minusDays = minusDays < 1 ? 1 : minusDays;
		// 檢查的天數
		return ++minusDays;
	}

	/**
	 * 取今日日期
	 * 
	 * @param delimiter
	 * @return
	 */
	protected static String getToDay () {
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		return sdf.format(cal.getTime());
		
	}

	/**
	 * 不足兩碼則前面補0
	 * 
	 * @param str
	 * @return
	 */
	protected static String getLen2(String str) {
		return str.length() < 2 ? "0" + str : str;
	}
	
	/**
     * 不為空
     */
	protected static boolean isntBlank(Cell cell) {
		return cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK;
	}
	
	/**
	 * 將失敗的job相關資訊寫入file中
	 * 
	 * @param path
	 */
	protected static void writeListFtoFile(String path, String str, boolean end) {
	    String destFile = path + "/JobF.txt";
	    FileOutputStream fos = null ;
	    FileInputStream fis = null ;
	    PrintWriter pw = null;
        byte[] buffer=new byte[10240];
	    int s;
		str = "\r\n\r\n ====== " + getToDay() + " " + str;

	    try {
	    	File f = new File(destFile);
	    	
	    	/**
	    	 * createNewFile
	    	 * true: 表示檔案不存在，並會自動產生檔案
	    	 * false: 表示檔案已存在
	    	 */
	    	if(f.createNewFile())
	    		System.out.println("已自動產生檔案");
	    	
	    	// 將原檔案內容讀出後與整併進要寫入的內容中(原內容放最後)
	    	fis = new FileInputStream(f);
	    	while((s = fis.read(buffer)) != -1) {
	    		str += new String(buffer,0,s);
	    	}
	    	
			if(end)
				str += "\r\n\r\n --------------------- END ----------------------- \r\n\r\n";
			
	    	// 將整理好的內容寫入檔案內
	    	fos = new FileOutputStream(f); // 第二參數設定是保留原有內容(預設false會刪)
			fos.write(str.getBytes());
			
			fos.flush();
			// 若要設定編碼則需透過OutputStreamWriter
			pw = new PrintWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));
		} catch (Exception ex) {
			System.out.println("== writeListFtoTXT Exception ==> " + ex.getMessage());
		} finally {
			try {
				fos.close();
				fis.close();
				pw.close();
			} catch (IOException e) {
				System.out.println("== writeListFtoTXT Finally Exception ==> " + e.getMessage());
			}
		}
	}
}
