package com.gss.RunDailyReport;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * COPY來的
 * 
 * @author https://codertw.com/%E7%A8%8B%E5%BC%8F%E8%AA%9E%E8%A8%80/313681/
 */
public class Zip {

//	public static void main(String[] args) throws IOException {
//		/**
//		 * 壓縮檔案
//		 */
//		File[] files = new File[] { new File("d:/English"), new File("d:/發放資料.xls"), new File("d:/中文名稱.xls") };
//		File zip = new File("d:/壓縮.zip");
//		ZipFiles(zip, "abc", files);
//		/**
//		 * 解壓檔案
//		 */
//		File zipFile = new File("D:\\Downloads\\QC(Log)_406624.zip");
//		String path = "D:\\Downloads\\(406624)";
//		unZipFiles(zipFile, path);
//	}

	/**
	 * 壓縮檔案-由於out要在遞迴呼叫外,所以封裝一個方法用來
	 * 呼叫ZipFiles(ZipOutputStream out,String path,File... srcFiles)
	 * 
	 * @param zip
	 * @param path
	 * @param srcFiles
	 * @throws IOException
	 * @author isea533
	 */
	protected static void ZipFiles(File zip, String path, File... srcFiles) throws IOException {
		ZipOutputStream out = new ZipOutputStream(new FileOutputStream(zip));
		ZipFiles(out, path, srcFiles);
		out.close();
		System.out.println("*****************壓縮完畢*******************");
	}

	/**
	 * 壓縮檔案-File
	 * 
	 * @param zipFile  zip檔案
	 * @param srcFiles 被壓縮原始檔
	 * @author isea533
	 */
	protected static void ZipFiles(ZipOutputStream out, String path, File... srcFiles) {
		path = path.replaceAll("\\*", "/");
		if (!path.endsWith("/")) {
			path = "/";
		}
		byte[] buf = new byte[1024];
		try {
			for (int i = 0; i < srcFiles.length; i++) {
				if (srcFiles[i].isDirectory()) {
					File[] files = srcFiles[i].listFiles();
					String srcPath = srcFiles[i].getName();
					srcPath = srcPath.replaceAll("\\*", "/");
					if (!srcPath.endsWith("/")) {
						srcPath = "/";
					}
					out.putNextEntry(new ZipEntry(path + srcPath));
					ZipFiles(out, path + srcPath, files);
				} else {
					FileInputStream in = new FileInputStream(srcFiles[i]);
					System.out.println(path + srcFiles[i].getName());
					out.putNextEntry(new ZipEntry(path + srcFiles[i].getName()));
					int len;
					while ((len = in.read(buf)) > 0) {
						out.write(buf, 0, len);
					}
					out.closeEntry();
					in.close();
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 解壓檔案到指定目錄
	 * 
	 * @param zipFile
	 * @param descDir
	 * @author isea533
	 */
	@SuppressWarnings("rawtypes")
	protected static void unZipFiles(File zipFile, String descDir) throws IOException {

		ZipFile zip = null;
		try {
			File pathFile = new File(descDir);
			if (!pathFile.exists()) {
				pathFile.mkdirs();
			}
			zip = new ZipFile(zipFile);
			for (Enumeration entries = zip.entries(); entries.hasMoreElements();) {
				ZipEntry entry = (ZipEntry) entries.nextElement();
				String zipEntryName = entry.getName();
				InputStream in = zip.getInputStream(entry);
				String outPath = (descDir + File.separator + zipEntryName).replaceAll("\\\\", "/");

				// 判斷路徑是否存在,不存在則建立檔案路徑
				File file = new File(outPath.substring(0, outPath.lastIndexOf('/')));
				if (!file.exists()) {
					file.mkdirs();
				}
				// 判斷檔案全路徑是否為資料夾,如果是上面已經上傳,不需要解壓
				if (new File(outPath).isDirectory()) {
					continue;
				}
				// 輸出檔案路徑資訊
				System.out.println(outPath);
				OutputStream out = new FileOutputStream(outPath);
				byte[] buf1 = new byte[1024];
				int len;
				while ((len = in.read(buf1)) > 0) {
					out.write(buf1, 0, len);
				}
				in.close();
				out.close();
			}
			zip.close();
			System.out.println("******************解壓完畢********************");
		} catch (Exception ex) {
			System.out.println("Zip unZipFiles Error:");
			ex.printStackTrace();
		} finally {
			zip.close();
		}
	}

}
