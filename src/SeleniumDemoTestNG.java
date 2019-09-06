import org.testng.annotations.Test;
import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.AfterTest;
import org.testng.annotations.AfterSuite;

import java.awt.AWTException;
import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
//import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.google.common.io.Files;

//import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class SeleniumDemoTestNG {

	static int rowCount;
	static String localDirPath = "C:\\Users\\rpathania\\Desktop\\Selenium_files\\";
	static String outLogFolder="Log_files\\";
	static String fileName = "Interfaces_testing_copy.xlsx";
	static String sheetName = "test_intf";
	static int rowStartIndex = 1;
	static String logFileName;
	static String runCntrlId;

	public static void main(String[] args) throws InterruptedException, IOException, AWTException {

		File file = new File(localDirPath + fileName);
		
		JFrame f = new JFrame();

		String name = JOptionPane.showInputDialog(f,
				"Please select the Instance: \r\n 1>DSERAJ \r\n 2>DSER1J \r\n 3>SSERAJ \r\n 4>OCI");
		System.out.println(name);

		// System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")
		// + "\\chromedriver_win32_new\\chromedriver.exe");
		//System.setProperty("webdriver.chrome.driver", "C:\\chromedriver1.exe");
		System.setProperty("webdriver.gecko.driver", "C:\\geckodriver.exe");
		//WebDriver driver = new ChromeDriver();
		WebDriver driver = new FirefoxDriver();
		WebDriverWait wait = new WebDriverWait(driver, 50);

		try {
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
			SeleniumDemoTestNG ob = new SeleniumDemoTestNG();
			// Login
			String instName = ob.login(driver, Integer.parseInt(name));
			// Login
			String[][] res = new String[10][1];
			res = ob.navigateToPage(driver);

			for (int row = rowStartIndex - 1; row < rowCount; row++) {

				String[] navigation = res[row][0].split(">");
				String[] Arr_Prcsname = res[row][1].split(">");
				String Prcs = Arr_Prcsname[0].trim();
				String FTP_Folder = Arr_Prcsname[1].trim();
				String fileName = Arr_Prcsname[2].trim();
				
				fileName = fileName.replace("DD", "%dd%");
				fileName = fileName.replace("MM", "%mm%");
				fileName = fileName.replace("YYYY", "%yy%");
				
				if (fileName.equals("NA"))
				fileName = "*.*";
				
				System.out.println("File Name = "+fileName);

				System.out.println(navigation + " " + Prcs);
				try {

				for (int i = 0; i < navigation.length; i++) {
					Thread.sleep(500);
					navigation[i] = navigation[i].trim();
					WebElement myElement = driver.findElement(By.linkText(navigation[i]));
					myElement.click();
				}
				}
				catch(NoSuchElementException ex1)
				{
					System.out.println("Now what?");
				}

				ob.addPage(driver, wait, Arr_Prcsname[0].trim());
				String stat = ob.prcsMonitor(driver, wait, Prcs, row + 1);

				Thread.sleep(3000);
				if (!FTP_Folder.equals("NA") && stat.equals("Success")) {
					ob.createBat(Prcs, Prcs, FTP_Folder, instName,fileName);
					ob.runBat(Prcs);
					ob.checkFTP(Prcs, row + 1);

				}
				driver.navigate().refresh();

			}

			driver.close();
			Desktop desktop = Desktop.getDesktop();
			if (file.exists())
				desktop.open(file);

		} 
		
		catch (Exception ex) {
			ex.printStackTrace();
			driver.close();

		}

	}

	public String login(WebDriver driver, int instance) throws IOException {

		String URL;
		String userid;
		String password;
		String instName;
		//String runCntrlId;
		//String sheetToTest;
		File file = new File(localDirPath + fileName);
		FileInputStream inputStream = new FileInputStream(file);
		Workbook readWorkbook = null;
		readWorkbook = new XSSFWorkbook(inputStream);
		Sheet readSheet = readWorkbook.getSheet("login");
		// rowCount = readSheet.getLastRowNum() - readSheet.getFirstRowNum();
		/*
		 * for (int i = 1; i < rowCount + 1; i += 1) {
		 * 
		 * 
		 * 
		 * 
		 * if (instance.equals(row.getCell(0).getStringCellValue().trim())) {
		 * 
		 * 
		 * break;
		 * 
		 * } else continue;
		 * 
		 * 
		 * }
		 */
		Row row = readSheet.getRow(instance);
		instName = row.getCell(0).getStringCellValue().trim();
		URL = row.getCell(1).getStringCellValue();
		System.out.println(URL);
		userid = row.getCell(2).getStringCellValue();
		password = row.getCell(3).getStringCellValue();
		runCntrlId = row.getCell(4).getStringCellValue();
		sheetName = row.getCell(5).getStringCellValue().trim();
		System.out.println("Sheet:"+sheetName);

		// driver.get("https://hris-steria-trn.oracleoutsourcing.com/psp/DSERAJ/?cmd=login&languageCd=ENG");
		driver.get(URL);

		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

		WebElement username = driver.findElement(By.id("userid"));
		// username.sendKeys("PS");
		username.sendKeys(userid);

		WebElement Password = driver.findElement(By.id("pwd"));
		// Password.sendKeys("test@123");
		Password.sendKeys(password);

		WebElement loginbutton = driver.findElement(By.name("Submit"));
		loginbutton.click();
		inputStream.close();
		readWorkbook.close();
		return instName;

	}

	public String[][] navigateToPage(WebDriver driver) throws IOException {

		String[][] nav = new String[50][2];
		File file = new File(localDirPath + fileName);
		FileInputStream inputStream = new FileInputStream(file);
		Workbook readWorkbook = null;
		readWorkbook = new XSSFWorkbook(inputStream);
		Sheet readSheet = readWorkbook.getSheet(sheetName);
		rowCount = readSheet.getLastRowNum() - readSheet.getFirstRowNum();
		
		
		for (int j = 1; j < rowCount + 1; j += 1) {
			Row row = readSheet.getRow(j);
			Cell cell = row.getCell(6);
			if (cell == null || cell.getStringCellValue().trim().isEmpty())
			{
				//System.out.println(cell.getStringCellValue());
				rowStartIndex = j;
				break;
			}
				
			else 
				continue;
				
			
		}
		System.out.println("j=" + rowStartIndex);

		for (int i = rowStartIndex; i < rowCount + 1; i += 1) {

			Row row = readSheet.getRow(i);

			// Create a loop to print cell values in a row
			System.out.println("rowcount=" + rowCount + " " + rowStartIndex);
			nav[i - 1][0] = row.getCell(3).getStringCellValue();
			nav[i - 1][1] = row.getCell(4).getStringCellValue();


		}
		readWorkbook.close();
		return nav;

	};

	public void addPage(WebDriver driver, WebDriverWait wait, String Prcsname) throws InterruptedException {

		/*
		 * String RunId = (new
		 * java.sql.Timestamp(System.currentTimeMillis())).toString().replaceAll(":","")
		 * .replaceAll(" ",""); RunId= RunId.substring(5,RunId.length());
		 */

		//Thread.sleep(5000);
		wait.until(ExpectedConditions.frameToBeAvailableAndSwitchToIt("ptifrmtgtframe"));
		WebElement Srch_PS = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("PSEDITBOX")));
		Srch_PS.sendKeys(runCntrlId);

		WebElement Add_New = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("#ICSearch")));
		Add_New.click();
		WebElement Run_Intf = wait
				.until(ExpectedConditions.visibilityOfElementLocated(By.name("PRCSRQSTDLG_WRK_LOADPRCSRQSTDLGPB")));
		Thread.sleep(1000);
		Run_Intf.click();
		Thread.sleep(1000);
		WebElement ChckBox = wait.until(
				ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(text(),'" + Prcsname + "')]")));
		String x = ChckBox.getAttribute("id");
		String x1 = x.substring(x.lastIndexOf("$") + 1);
		System.out.println(x1);
			WebElement ChckBox1 = wait
				.until(ExpectedConditions.visibilityOfElementLocated(By.id("PRCSRQSTDLG_WRK_SELECT_FLAG$" + x1)));
		if (!ChckBox1.isSelected()) {
			ChckBox1.click();
		}
		Thread.sleep(2000);
		WebElement OK = driver.findElement(By.id("#ICSave"));
		OK.click();

	}

	public String prcsMonitor(WebDriver driver, WebDriverWait wait, String Prc, int Rw)
			throws InterruptedException, IOException {
		/* Process Monitor code start */
		// int retCode =0;
		String stat;
		FileWriter fw;
		String PrIn;
		
		
		
		WebElement PrcsInstnum = wait
				.until(ExpectedConditions.visibilityOfElementLocated(By.id("win0divPRCSRQSTDLG_WRK_DESCR100")));
		PrIn = PrcsInstnum.getText().substring(17);

		WebElement PrcsMonitorLink = wait
				.until(ExpectedConditions.visibilityOfElementLocated(By.name("PRCSRQSTDLG_WRK_LOADPRCSMONITORPB")));
		PrcsMonitorLink.click();
		Thread.sleep(5000);
		WebElement PrcInstNumField = driver.findElement(By.id("PMN_DERIVED_PRCSINSTANCE"));
		PrcInstNumField.clear();
		Thread.sleep(3000);
		PrcInstNumField.sendKeys(PrIn);

		WebElement Refresh = driver.findElement(By.id("REFRESH_BTN"));
		Refresh.click();
		Thread.sleep(1000);
		fw = new FileWriter(localDirPath + outLogFolder +"logs\\"+ Prc + ".txt");
		WebElement InstanceNumber = driver.findElement(By.id("PMN_PRCSLIST_PRCSINSTANCE$0"));
		String InstNum = InstanceNumber.getText();
		fw.write("Process Instance number: " + InstNum + System.getProperty("line.separator"));
		boolean flag = true;
		// Thread.sleep(9000);

		do {
			WebElement Status = driver.findElement(By.id("PMN_PRCSLIST_RUNSTATUSDESCR$0"));
			System.out.println(Status.getText());
			stat = Status.getText();
			fw.write("\n Run Status: " + stat + System.getProperty("line.separator"));
			if (stat.equals("Processing") || stat.equals("Queued") || stat.equals("Initiated")
					|| stat.equals("Pending")) {
				Thread.sleep(10000);
				WebElement Rfrsh = driver.findElement(By.id("REFRESH_BTN"));
				Rfrsh.click();
			} else if (stat.equals("Success") || stat.equals("No Success") || stat.equals("Error")) {
				System.out.println("The process ended with " + stat);
				WebElement DistStatus = driver.findElement(By.id("PMN_PRCSLIST_DISTSTATUS$0"));
				System.out.println(DistStatus.getText());
				fw.write("\n Run Status: " + stat + " Dist Status: " + DistStatus.getText()
						+ System.getProperty("line.separator"));
				if (!DistStatus.getText().equals("Posted")) {
					WebElement Rfrsh = driver.findElement(By.id("REFRESH_BTN"));
					Rfrsh.click();
					Thread.sleep(3000);
					continue;
				}

				flag = false;
			}

			else if (stat.equals("Cancelled")) {

				System.out.println("Process is cancelled");

				fw.write("\n Process is cancelled" + System.getProperty("line.separator"));
				fw.close();
				// break;
				// System.exit(1);
			}

		} while (flag);
		Thread.sleep(3000);

		fw.close();

		/**/
		// File file = new File("D:\\Profiles\\rpathania.EMEAAD\\Desktop\\Nav.xlsx");
		File file = new File(localDirPath + fileName);
		FileInputStream inputStream = new FileInputStream(file);
		Workbook readWorkbook = null;
		readWorkbook = new XSSFWorkbook(inputStream);
		Sheet readSheet = readWorkbook.getSheet(sheetName);
		// rowCount = readSheet.getLastRowNum() - readSheet.getFirstRowNum();

		// for (int i = 0; i < rowCount+1; i+=1) {

		Row row = readSheet.getRow(Rw);

		// Create a loop to print cell values in a row

		// nav= row.getCell(i).getStringCellValue();
		// nav[i-1][0] = row.getCell(3).getStringCellValue();
		// nav[i-1][1] = row.getCell(4).getStringCellValue();

		// for (int j = 0; j < row.getLastCellNum(); j++) {

		// Print Excel data in console

		// System.out.print(row.getCell(j)+"|| ");

		// }
		Cell cell5 = row.createCell(5);// Process Instance cell
		cell5.setCellValue(PrIn);
		Cell cell7 = row.createCell(7);// FTP Instance cell
		cell7.setCellValue("N/A");
		CellStyle style = readWorkbook.createCellStyle();
		CellStyle style1 = readWorkbook.createCellStyle();

		style.setFillForegroundColor(IndexedColors.RED1.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style1.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		// Cell cell = row.createCell(row.getLastCellNum());
		Cell cell6 = row.createCell(6); // test passed/failed column

		System.out.println("row no=" + Rw);
		if (Rw == 0) {
			cell6.setCellValue("Status");
			cell6.setCellStyle(style);
		}
		
		if (stat.equals("Success")) {
			// cell.setCellStyle(style1);
			cell6.setCellValue("Passed");
			cell6.setCellStyle(style1);
		} else {
			// cell.setCellStyle(style);
			cell6.setCellValue("Failed");
			cell6.setCellStyle(style);
		}
		inputStream.close();

		FileOutputStream outputStream = new FileOutputStream(file);

		readWorkbook.write(outputStream);
		outputStream.close();

//    }
		readWorkbook.close();
		return stat;
		
		/* Process Monitor code ends */
	}

	public void createBat(String logFileName, String batFileName, String FTP_Folder, String Instance,String fileName)
			throws IOException {
		String dt = "set yy=%date:~-4%\r\n" + "set mm=%date:~-10,2%\r\n" + "set dd=%date:~-7,2%\r\n"
				+ "set MYDATE=%yy%%mm%%dd%\r\n";
		String s1 = "echo. >> " + localDirPath + outLogFolder +"logs\\"+  logFileName + ".txt \n";

		String s2 = "pscp -i C:\\pkey.ppk  i_sera@132.240.135.110:/" + Instance + "/outgoing/" + FTP_Folder
				+ "/"+fileName +" " +localDirPath + outLogFolder+"files\\ >> " + localDirPath + outLogFolder
				+ "logs\\"+logFileName + ".txt \n";
		File BatFile = new File(localDirPath + outLogFolder + "logs\\"+batFileName + ".bat");
		FileOutputStream fos = new FileOutputStream(BatFile);
		DataOutputStream dos = new DataOutputStream(fos);
		dos.writeBytes(dt);
		dos.writeBytes(s1);
		dos.writeBytes(s2);
		dos.close();
	}

	public void runBat(String batName) throws InterruptedException, AWTException, IOException {
		Thread.sleep(2000);

		ProcessBuilder processBuilder = new ProcessBuilder(localDirPath + outLogFolder +"logs\\"+batName + ".bat");

		try {

			Process process = processBuilder.start();

			StringBuilder output = new StringBuilder();

			BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()));

			String line;
			while ((line = reader.readLine()) != null) {
				output.append(line + "\n");
			}

			int exitVal = process.waitFor();
			if (exitVal == 0) {
				System.out.println(output);
				// System.exit(0);
			} else {
				// abnormal...
			}

		} catch (IOException e) {
			e.printStackTrace();
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	public void checkFTP(String Prc, int Rw) {
		try {
			FileReader reader = new FileReader(localDirPath + outLogFolder   +"logs\\"+Prc + ".txt");
			BufferedReader bufferedReader = new BufferedReader(reader);

			String line;
			boolean flag1 = false;

			while ((line = bufferedReader.readLine()) != null) {
				System.out.println(line);
				if (line.contains("ETA:")) {
					System.out.println("Yes");
					flag1 = true;
					break;

				}
			}
			File file = new File(localDirPath + fileName);
			FileInputStream inputStream = new FileInputStream(file);
			Workbook readWorkbook = null;
			readWorkbook = new XSSFWorkbook(inputStream);
			Sheet readSheet = readWorkbook.getSheet(sheetName);

			Row row = readSheet.getRow(Rw);
			System.out.println(row.getCell(3).getStringCellValue());
			Cell cell7 = row.createCell(7);// FTP status column
			CellStyle style = readWorkbook.createCellStyle();
			CellStyle style1 = readWorkbook.createCellStyle();
			style.setFillForegroundColor(IndexedColors.RED1.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style1.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
			style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			if (flag1) {
				System.out.println("Pass");
				cell7.setCellValue("Passed");
				cell7.setCellStyle(style1);
			} else {
				System.out.println("Failed");
				cell7.setCellValue("Failed");
			    cell7.setCellStyle(style);
			}

			reader.close();
			inputStream.close();

			FileOutputStream outputStream = new FileOutputStream(file);

			readWorkbook.write(outputStream);
			outputStream.close();

			readWorkbook.close();

		} catch (IOException e) {
			e.printStackTrace();
			// readWorkbook.close();
		}
	}
}
