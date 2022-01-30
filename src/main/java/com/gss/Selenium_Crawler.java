package com.gss;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.jsoup.helper.StringUtil;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;

import us.codecraft.webmagic.selector.Html;

public class Selenium_Crawler {

	static final String mailStartText = " - 您好, 〔("; // job 寄的 mail 開頭
	static WebDriver driver = null;
	static Html html;
	static WebElement element;

	protected static List<Map<String, String>> getMailContent(
			String path, String[] inboxName, String account,
			String pwd, Calendar cal, ArrayList<TreeMap<String, String>> listFforSheet3) {
		
		driver = null;
		List<WebElement> listElement;
		List<Map<String, String>> listMap = new ArrayList<Map<String, String>>();
		List<Map<String, String>> qcLogList;
		Map<String, String> map;
		Map<String, String> qcLog;
		File zipFile;
		boolean download = false;
		String body = "", title = "", jobRSDate = "", job_id = "", job_seq = "", unZipFile = "";
		String[] jobMailTitleArr;
		int arrLen = 0;

		String os = System.getProperty("os.name");
		Map<String, String> mapProp = Property.getProperties(path);
		Integer chkDate = Integer.valueOf(mapProp.get("chkDate"));
		String chromeDefaultDownloadPath = os.contains("Mac")
				? mapProp.get("chromeDefaultDownloadPathMac")
				: mapProp.get("chromeDefaultDownloadPathWindows");

		// Selenium
		DesiredCapabilities capabilities = DesiredCapabilities.chrome();
		capabilities.setCapability("chrome.switches", Arrays.asList("--start-maximized"));
		System.setProperty("webdriver.chrome.driver",
				path + mapProp.get("chromedriver") + (os.contains("Mac") ? "" : ".exe"));
		driver = new ChromeDriver(capabilities);
		driver.get(mapProp.get("webAddress"));
		System.out.println("##start login ");

		try {
			// 延迟加载，保证JS数据正常加载
			Thread.sleep(1000);

			// 登入Mail
			driver.findElement(By.id("username")).sendKeys(account);
			driver.findElement(By.id("password")).sendKeys(pwd);
			driver.findElement(By.className("DwtButton")).click();

			// 等待三秒以確保頁面加載完整
			Thread.sleep(3000);
			
			for (String str : inboxName) {
				qcLogList = new ArrayList<Map<String, String>>();

				// 進入對應的收信匣
				listElement = driver.findElements(By.xpath("//td[@class='DwtTreeItem-Text']"));
				for (WebElement em : listElement) {
					if (em.getText().contains(str)) {
						System.out.println(em.getText());
						download = str.contains("失敗");
						/**
						 * 因使用原em.click();
						 * 常造成點擊第二個信箱項目時失敗
						 * 失敗原因可能為driver版本與實際chrome版本不合
						 * (即使是透過舊版chrome driver開啟也一樣會被自動更新為新版)
						 * 故改用Actions取代
						 */
//						em.click();
						new Actions(driver).moveToElement(em).click().perform();
						break;
					}
				}

				Thread.sleep(1000);
				/**
				 * 內容加載後截取信件list區塊 再拆分為主旨、內容兩部份放入map中
				 */
				if (scrollDown(cal)) {
					element = driver.findElement(By.id("zl__TV-main__rows"));
					html = new Html(element.getAttribute("outerHTML"));
					listElement = driver.findElements(By.className("Row"));

					for (WebElement em : listElement) {
						map = new HashMap<String, String>();
						title = em.getAttribute("aria-label");
						map.put("title", title);

						// 使用getText()會發生明明有資料卻讀不到的情況
						// 因此改用.getAttribute("innerHTML")
						body = em.findElement(By.className("ZmConvListFragment")).getAttribute("innerHTML");
						// 排除非p開頭的job_id
						if (body.indexOf(")p_") < 0)
							continue;
						map.put("body", body);
						listMap.add(map);

						/**
						 * jobMailTitleArr 最後一個逗號有時會多空隔有時不會，因此先將空格移除再做split
						 * 0: 已讀
						 * 1: NHIA
						 * 2: 有附件
						 * 3: 執行成功=>(386062)彙總－疾病就醫利用彙整檔(補歷史資料)(DWF)_2007/01(附檔)
						 * 4: 收件匣/NHIA/2JOB成功
						 * 5: 18KB
						 * 6: "今天" or "21/1/16六"
						 * 7: 上午10:22
						 */
						jobMailTitleArr = title.replace(" ", "").split(",");
						// 有時job的中文內容會有逗號，會影響到陣列的總數
						arrLen = jobMailTitleArr.length;
						// job 所屬日期 (日誌用)
						jobRSDate = jobMailTitleArr[arrLen - 2].trim();
						job_seq = body.substring(body.indexOf("〔(") + 2);
						job_seq = job_seq.substring(0, job_seq.indexOf(")p_"));
						job_id = body.substring(body.indexOf(")p_") + 1);
						job_id = job_id.substring(0, job_id.indexOf("-"));

						/**
						 * 下載附件的條件
						 * 1 在失敗的資料夾內
						 * 2 此mail的內容為job執行結果的mail
						 * 3 收信時間在欲檢查的時間內
						 */
						if (download && body.indexOf(mailStartText) == 0 && isDownloadFile(jobRSDate, chkDate - 1)) {

							em.click();
							Thread.sleep(1000);
							element = driver.findElement(By.id("zv__TV__TV-main_MSG_attLinks_2_main"));
							element.click();
							Thread.sleep(2000);

							title = element.getAttribute("title");
							/**
							 * 解壓檔案
							 */
							zipFile = new File(chromeDefaultDownloadPath + title);
							unZipFile = "(" + job_seq + ")" + job_id;
							Zip.unZipFiles(zipFile, chromeDefaultDownloadPath + unZipFile);

							// 整理RQLog的相關資訊
							qcLog = new HashMap<String, String>();
							qcLog.put("qcLogFile", unZipFile);
							qcLog.put("qcLogExcel", "Log_" + job_seq + ".xls");
							qcLogList.add(qcLog);
						}
					}

					/**
					 * 將RQ的狀態回寫到對應LOG檔檔名上
					 * 整理出失敗的job資訊寫入jobF.txt
					 */
					if(download)
						LogRename.logRename(path, qcLogList, chromeDefaultDownloadPath, listFforSheet3);
				}
			}

		} catch (Exception e) {
			System.out.println("Selenium_Crawler Error：");
			e.printStackTrace();
		} finally {
			driver.close();
		}
		return listMap;
	}

	/**
	 * 是否下載附件檔案
	 * 
	 * @param title
	 * @param chkDate
	 * @return
	 */
	private static boolean isDownloadFile(String jobRSDate, Integer chkDate) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
		String yy1 = "20"; // 西元年前兩碼

		// 取 昨、今、明 三天的日期
		Calendar cal = Calendar.getInstance();
		String today = sdf.format(cal.getTime());
		cal.add(Calendar.DATE, -1);
		String yesterday = sdf.format(cal.getTime());

		// job 原日期 (後面刪除相同job時使用)
		String jobRSOriDate = jobRSDate.equals("昨天") ? yesterday : jobRSDate.equals("今天") ? today : "";
		if (jobRSDate.lastIndexOf("/") > 0) {
			String[] jobRSDateArr = jobRSDate.split("/");
			String yy2 = jobRSDateArr[0];
			String mm = jobRSDateArr[1];
			String dd = jobRSDateArr[2].substring(0, jobRSDateArr[2].length() - 1);
			jobRSOriDate = yy1 + yy2 + Tools.getLen2(mm) + Tools.getLen2(dd);
		}
		return Integer.valueOf(jobRSOriDate) >= chkDate;
	}

	/**
	 * 滑動頁面
	 */
	private static boolean scrollDown(Calendar cal) {
		if (driver != null) {
			try {
				SimpleDateFormat sdfyyMd = new SimpleDateFormat("yy/M/d");
				element = driver.findElement(By.id("zl__TV-main__rows"));
				html = new Html(element.getAttribute("outerHTML"));
				// 滑動頁面直到信件時間出現欲檢查日期的前一天出現為止
				String scroll = "go";
				int chkMailDateLen = 5;
				String[] calArr = new String[chkMailDateLen];
				for (int i = 0; i < chkMailDateLen; i++) {
					calArr[i] = sdfyyMd.format(cal.getTime());
					cal.add(Calendar.DATE, -1);
				}
				cal.add(Calendar.DATE, +5);
				while ("go".equals(scroll)
					&& StringUtil.isBlank(html.xpath("//li[contains(@aria-label,', " + calArr[0] + "')]").get())
					&& StringUtil.isBlank(html.xpath("//li[contains(@aria-label,', " + calArr[1] + "')]").get())
					&& StringUtil.isBlank(html.xpath("//li[contains(@aria-label,', " + calArr[2] + "')]").get())
					&& StringUtil.isBlank(html.xpath("//li[contains(@aria-label,', " + calArr[3] + "')]").get())
					&& StringUtil.isBlank(html.xpath("//li[contains(@aria-label,', " + calArr[4] + "')]").get())
					) {

					if ("go".equals(scroll)) {
						html = new Html(element.getAttribute("outerHTML"));
						// 執行頁面滾動的JS語法
						String height1 = ((JavascriptExecutor) driver)
								.executeScript("var element = document.getElementById('zl__TV-main__rows');"
										+ "var height1 = element.scrollHeight;"
										+ "element.scroll(0,height1);"
										+ "return height1;")
								.toString();
						Thread.sleep(1000);
						String height2 = ((JavascriptExecutor) driver)
								.executeScript("var element = document.getElementById('zl__TV-main__rows');"
										+ "var height2 = element.scrollHeight;"
										+ "return height2;")
								.toString();
						scroll = Integer.parseInt(height1) == Integer.parseInt(height2) ? "stop" : "go";
						// 给页面预留加载时间
						Thread.sleep(1000);
					}
				}
				System.out.println("加載中...");
				return true;
			} catch (Exception e) {
				System.out.println("加載失敗:");
				e.printStackTrace();
			}
		}
		return false;
	}

}
