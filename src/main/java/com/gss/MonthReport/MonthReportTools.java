package com.gss.MonthReport;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;

public class MonthReportTools {
	
	private static long diffrence;
	private static final TimeUnit time = TimeUnit.MINUTES;
	private static final List<String> listIssueType1 = Arrays.asList(new String[] { "010" });
	private static final List<String> listIssueType2 = Arrays.asList(new String[] { "011", "012" });
	private static final List<String> listIssueType3 = Arrays
			.asList(new String[] { "013", "014", "015", "018", "019" });
	/**
	 * 驗証結果
	 * 
	 * @param acceptDate
	 * @param replyDate
	 * @param dueDate
	 * @param actDate
	 * @param issueType
	 * @param manualChk
	 * @return
	 */
	@SuppressWarnings("deprecation")
	protected static String getValidResult(Date acceptDate, Date replyDate, Date dueDate, Date actDate,	String issueType) {
		String validResult = "";
		/**
		 * 驗證內容是否有誤
		 * 回應時間 - 受理時間 需大於等於 1
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
			} else if (diffrence <= 0) {
				validResult = "ERR-回應時間 - 受理時間 需大於等於 1";
			} else if (actDate != null) {
				diffrence = time.convert(actDate.getTime() - dueDate.getTime(), TimeUnit.MILLISECONDS);
				if (diffrence / 60f / 24f >= 1) {
					validResult = "ERR-到期日不可早於實際完成日";
				} else
					validResult = "Normal";
			}else
				validResult = "Normal";
		}

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
	protected static boolean isFutureDate(Date actDate) {
		Calendar c = Calendar.getInstance();
		c.setTime(actDate);
		if (c.get(Calendar.MONTH) >= Calendar.getInstance().get(Calendar.MONTH))
			return true;
		return false;
	}
	
	/**
	 * 是否為過去日期
	 * 
	 * @param actDate
	 * @return
	 */
	protected static boolean isPastDate(Date acceptDate) {
		Calendar initReportDate = Calendar.getInstance();
		initReportDate.add(Calendar.MONTH, -2);
		Calendar c = Calendar.getInstance();
		c.setTime(acceptDate);
//System.out.println("c:"+c.get(Calendar.YEAR)+"/"+(c.get(Calendar.MONTH)+1)+"/"+c.get(Calendar.DATE));
//System.out.println("initReportDate:"+initReportDate.get(Calendar.YEAR)+"/"+(initReportDate.get(Calendar.MONTH)+1)+"/"+initReportDate.get(Calendar.DATE));
		if (c.get(Calendar.MONTH) < initReportDate.get(Calendar.MONTH))
			return true;
		return false;
	}
	
	/**
	 * 取得系統月份 - 1
	 * 
	 * @param actDate
	 * @return
	 */
	protected static int getLastMonth() {
		Calendar c = Calendar.getInstance();
		return c.get(Calendar.MONTH);
	}

	/**
	 * 因日期欄位的TYPE有時會變String有時會是Numeric
	 * 
	 * @param actDate
	 * @return
	 * @throws ParseException 
	 */
	protected static Date getDateValue(Cell cell, SimpleDateFormat sdf) throws ParseException {
		if(cell.getCellType() == Cell.CELL_TYPE_STRING)
			return sdf.parse(cell.getStringCellValue());
		if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
			return cell.getDateCellValue();
		
		return null;
	}
}
