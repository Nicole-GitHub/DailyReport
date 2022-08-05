package com.gss.RunDailyReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.Map.Entry;

import org.jsoup.helper.StringUtil;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.DesiredCapabilities;

import com.gss.Property;
import com.gss.Tools;

import us.codecraft.webmagic.selector.Html;

public class Selenium_Crawler {

	static final String mailStartText = " - 您好, 〔("; // job 寄的 mail 開頭
	static ChromeDriver driver = null;
	static Html html;
	static WebElement element;

	protected static List<Map<String, String>> getMailContent(
			String path, String[] inboxName, String account,
			String pwd, Calendar cal, ArrayList<TreeMap<String, String>> listFforSheet3) throws Exception {
		
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

		Calendar chkDate = cal;
		chkDate.add(Calendar.DATE, -1);
		Integer chkDateYesterdayInt = Integer.valueOf(Tools.getCalendar2String(chkDate, "yyyyMMdd"));

		String chromeDriverPath = mapProp.get("chromeDriverPath");
		String chromeDriverName = mapProp.get("chromeDriverName");
		String chromeDriverVersion = mapProp.get("chromeDriverVersion");
		String chromeDefaultDownloadPath = os.contains("Mac")
				? mapProp.get("chromeDefaultDownloadPathMac")
				: mapProp.get("chromeDefaultDownloadPathWindows");
				
		String chromeDriver = path + chromeDriverPath + chromeDriverName + "_" + chromeDriverVersion
				+ (os.contains("Mac") ? "" : ".exe");
		
		String unZipFilePath = path + "QC(Log)_" + Tools.getToDay("yyyyMMddHHmmss") + "/";
		
		// Selenium
		DesiredCapabilities capabilities = DesiredCapabilities.chrome();
		capabilities.setCapability("chrome.switches", Arrays.asList("--start-maximized"));
		System.setProperty("webdriver.chrome.driver", chromeDriver);

		driver = new ChromeDriver(capabilities);
		
		// 下載新的ChromeDriver
		String currentChromeVer = driver.getCapabilities().getVersion();
		currentChromeVer = currentChromeVer.substring(0,currentChromeVer.indexOf("."));
		if(!currentChromeVer.equals(chromeDriverVersion)) {
			downloadChromeDriver(path, driver, chromeDefaultDownloadPath , mapProp);
			// 改寫properties的chromeDriverVersion
			chgChromeDriverVersionFromProp(path, currentChromeVer, chromeDriverVersion);
		}
			
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
			int readInbox = 0;
			
			for (String str : inboxName) {
				qcLogList = new ArrayList<Map<String, String>>();

				// 進入對應的收信匣
				listElement = driver.findElements(By.xpath("//td[@class='DwtTreeItem-Text']"));
				for (WebElement em : listElement) {
					if (em.getText().contains(str)) {
						System.out.println(em.getText());
						// 失敗的job需下載附件
						download = str.contains("失敗");
//						em.click();
						readInbox++;
						/**
						 * 因使用原em.click();
						 * 常造成點擊第二個信箱項目時失敗
						 * 失敗原因可能為driver版本與實際chrome版本不合
						 * (即使是透過舊版chrome driver開啟也一樣會被自動更新為新版)
						 * 故改用Actions取代
						 */
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
						 * 1. 在失敗的資料夾內
						 * 2. 此mail的內容為job執行結果的mail
						 * 3. 收信時間在欲檢查的時間內
						 */
						if (download && body.indexOf(mailStartText) == 0 && isDownloadFile(jobRSDate, chkDateYesterdayInt)) {

							em.click();
							Thread.sleep(1000);
							element = driver.findElement(By.id("zv__TV__TV-main_MSG_attLinks_2_main"));
							element.click();
							Thread.sleep(2000);

							// 解壓檔案並放置對應目錄下
							zipFile = new File(chromeDefaultDownloadPath + element.getAttribute("title"));
							unZipFile = "(" + job_seq + ")" + job_id;
							Zip.unZipFiles(zipFile, unZipFilePath + unZipFile);
							
							// 刪除下載的檔案
							zipFile.delete();
							
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
						LogRename.logRename(path, qcLogList, unZipFilePath, listFforSheet3);
				}
			}
			
			// 若未讀取完整，則拋Exception
			if(readInbox < inboxName.length)
				throw new Exception("收件匣讀取數量過少!!");
			
			boolean isOK = false;
			for (Map<String, String> listmap : listMap) {
				if(listmap.get("title").contains("執行成功=>")) {
					isOK = true;
					break;
				}
			}
			
			if(!isOK)
				throw new Exception("收件匣讀取不完整!!");
			
		} catch (Exception e) {
			System.out.println("Selenium_Crawler Error：" + e.getMessage());
			throw e;
		} finally {
			if(driver != null)
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

		// 取 昨、今 兩天的日期
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
			jobRSOriDate = yy1 + yy2 + Tools.setLen(mm, 2) + Tools.setLen(dd, 2);
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
				String scroll = "go";
				/**
				 * 因不一定每天都有失敗job
				 * 故設定滑動頁面"最多"滑到信件時間出現欲檢查日期的前10天出現為止
				 */
				int chkMailDateLen = 10;
				
				String[] calArr = new String[chkMailDateLen];
				for (int i = 0; i < chkMailDateLen; i++) {
					calArr[i] = sdfyyMd.format(cal.getTime());
					cal.add(Calendar.DATE, -1);
				}

				// 因爬不只一個收件匣，故需將cal日期加回去
				cal.add(Calendar.DATE, chkMailDateLen);
				boolean dateisBlank = true,dateisReach = false;
				String checkStr = "", checkCal = "";
				
				while ("go".equals(scroll)) {
					// 最少要滾到檢查日期的前兩天,故初始值為1
					for (int calArrLen = 1; calArrLen < chkMailDateLen; calArrLen++) {
						checkCal = "//li[contains(@aria-label,', " + calArr[calArrLen] + " ')]";
						dateisBlank = StringUtil.isBlank(html.xpath(checkCal).get());
						// 若已滾到檢查日期的前兩天前則可停止
						if (!dateisBlank) {
							checkStr = html.xpath(checkCal).get();
							checkStr = checkStr.substring(checkStr.indexOf(","), checkStr.indexOf("<div"));
							checkStr = checkStr.substring(0, checkStr.lastIndexOf(","));
							checkStr = checkStr.substring(checkStr.lastIndexOf(",") + 1);
							checkStr = checkStr.substring(0, checkStr.lastIndexOf(" ")).trim();
							// 為避免出現22/6/11也被22/6/1 contains 到，故加此段確保日期正確
							if (checkStr.equals(calArr[calArrLen])) {
								dateisReach = true;
								break;
							}
						}
					}
					if (!dateisBlank && dateisReach)
						break;

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
					/**
					 * height1: 未滾前的高度
					 * height2: 滾動後的高度
					 * 若兩個高度皆相同則表示已滾到底
					 */
					scroll = Integer.parseInt(height1) == Integer.parseInt(height2) ? "stop" : "go";
					// 给页面预留加载时间
					Thread.sleep(1000);
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

	/**
	 * 下載新的ChromeDriver
	 * 
	 * @param path
	 * @param driver
	 * @param chromeDefaultDownloadPath
	 * @param mapProp
	 * @throws Exception
	 */
	private static void downloadChromeDriver(String path, ChromeDriver driver, String chromeDefaultDownloadPath,
			Map<String, String> mapProp) throws Exception {

		String method = "downloadChromeDriver";
		List<WebElement> listElement;
		List<Map<String,String>> listMap = new ArrayList<Map<String,String>>();
		Map<String,String> map ;
		String href = "", ver = "", unZipFilePath = "", unZipFileName = "", unZipFileNewName = "";
		boolean boo = false;
		File zipFile;
		
		String chromeDriverName = mapProp.get("chromeDriverName");
		List<Integer> arrs = Arrays.asList(new Integer[] {5,7}); // 5:mac64 ; 7:win32
		
		driver.get("https://chromedriver.chromium.org/downloads");
		System.out.println("##start " + method);
		
		try {
			// 延迟加载，保证JS数据正常加载
			Thread.sleep(1000);
			
			listElement = driver.findElements(By.cssSelector("#h\\.e02b498c978340a_87 > div > div > ul:nth-child(3) > li"));
			int index = 0;
			for (WebElement em : listElement) {
				// 只有前三行才是載點
				if(++index > 3)
					break;
				
				href = em.findElement(By.cssSelector("li:nth-child(" + index + ") > p > span.aw5Odc > a"))
						.getAttribute("href");
				ver = href.substring(href.lastIndexOf("=") + 1);
				ver = ver.substring(0,ver.indexOf("."));

				map = new HashMap<String,String>();
				map.put("href", href);
				map.put("ver", ver);
				listMap.add(map);
			}

			for (Map<String,String> forMap : listMap) {
				driver.get(forMap.get("href"));
				Thread.sleep(2000);
				for (Integer i : arrs) {
					
					/**
					 * 1. 判斷是否已有對應檔案
					 * 2. 下載檔案
					 * 3. 解壓檔案並放置對應目錄下
					 * 4. 更名
					 * 5. 刪除下載的檔案
					 */
					// 1
					zipFile = new File(chromeDefaultDownloadPath + "chromedriver_" + (i == 5 ? "mac64" : "win32") + ".zip");
					unZipFilePath = path + mapProp.get("chromeDriverPath");
					unZipFileName = unZipFilePath + chromeDriverName + (i == 5 ? "" : ".exe");
					unZipFileNewName = unZipFilePath + chromeDriverName + "_" + forMap.get("ver") + (i == 5 ? "" : ".exe");

					if (new File(unZipFileNewName).exists()) {
						boo = true;
						break;
					}
					
					// 2
					driver.findElement(
							By.cssSelector("body > table > tbody > tr:nth-child(" + i + ") > td:nth-child(2) > a"))
							.click();
					Thread.sleep(8000);

					// 3
					Zip.unZipFiles(zipFile, unZipFilePath);
					Thread.sleep(3000);

					new File(unZipFileName).renameTo(new File(unZipFileNewName)); // 4
					if(i == 5)
						new File(unZipFileNewName).setExecutable(true, false); // mac檔案需調整權限後才能使用
					
					zipFile.delete(); // 5
				}
				
				if(boo)
					break;
			}
			System.out.println("##END " + method);

		} catch (Exception e) {
			System.out.println(method + " Error：" + e.getMessage());
			throw e;
		}
	}

	/**
	 * 改寫properties的chromeDriverVersion
	 * 
	 * @param path
	 * @param currentChromeVer
	 * @throws Exception
	 */
	private static void chgChromeDriverVersionFromProp(String path, String currentChromeVer, String chromeDriverVersion) throws Exception {
		String method = "chgChromeDriverVersionFromProp";
		String destFile = path + "config.properties";
		String str = "" ;
		FileOutputStream fos = null;
		FileInputStream fis = null;
		PrintWriter pw = null;
		byte[] buffer = new byte[10240];
		int s;

		try {
			File f = new File(destFile);

			// 將原檔案內容讀出後修改currentChromeVer的值
			fis = new FileInputStream(f);
			while ((s = fis.read(buffer)) != -1) {
				str += new String(buffer, 0, s)
						.replace("chromeDriverVersion=" + chromeDriverVersion,
								"chromeDriverVersion=" + currentChromeVer);
			}

			// 將整理好的內容寫入檔案內
			fos = new FileOutputStream(f); // 第二參數設定是保留原有內容(預設false會刪)
			fos.write(str.getBytes());

			fos.flush();
			// 若要設定編碼則需透過OutputStreamWriter
			pw = new PrintWriter(new OutputStreamWriter(fos, StandardCharsets.UTF_8));
		} catch (Exception ex) {
			System.out.println(method + " Error：" + ex.getMessage());
			throw ex;
		} finally {
			try {
				fos.close();
				fis.close();
				pw.close();
			} catch (IOException e) {
				System.out.println(method + " Finally Error：" + e.getMessage());
				throw e;
			}
		}
	}
}
