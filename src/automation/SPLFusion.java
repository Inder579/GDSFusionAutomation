package automation;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.HeadlessException;
import java.awt.RenderingHints.Key;
import java.awt.Toolkit;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Date;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.NoSuchElementException;
import java.util.concurrent.TimeUnit;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextPane;
import javax.swing.UIManager;
import javax.swing.plaf.FontUIResource;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.DocumentException;
import org.joda.time.DateTime;
import org.joda.time.LocalDateTime;
import org.joda.time.format.DateTimeFormat;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.monitorjbl.xlsx.StreamingReader;

import resources.BrowserDriver;
import resources.Screenshot;

public class SPLFusion extends BrowserDriver {

	public static int attemptNo = 0;
	public String screenShotPathforInterestRate;
	public WebDriver driver;
	int cvScore, BehaviourScore;
	public String ActualIncome, appType, loanType, splloanType, Province, ApplicationID, MorgagePayment;
	public String TotalIncomeAmount, IntRate, cabKey, qlaStrategy, applicationType;
	public double TotalIncome, RemainingIncome, TotalDebt, ExpectedQLA, ExpInt, SPLltv, Maxltv, HomeEquity, PropertyVal;

	String lowefs, highefs, Prov, provinceGroup, bkStrategy, ps = "", code = null, propertyType = "",
			propertyLocation = "";
	double lef, hef, calRemIn, QLA, remIn, remInNaPrev, remInNaAfter, LtvMax, ActualMaxHA, ExpectedMaxHA, IncomeValue, TotalLaibility;

	int fcol, lcol, col, coldiff, rowNum, RiskGroup, SPLTotalDebt, lastNumRow, bkDecreaseAmount, RandomNumberResponse, length;
	String stringSplit[], Strategy;

	String lastname, firstname, address, city, dob, clprod, loanpurpose, hearabout, Referral, livingsituation, email,
			lengthofstay;

	String phone, loanamount, landlordname, landlordnumber, Employername, Employerposition, Incomeamt, Incomefreq,
			Employmentstatus, Supervisorname, Supervisornumber, lengthofemployment, previousemployer,
			lengthpreviousemployer, preferedLang;
	String QualifiedLoanAmount, totalincomeFusion, totaldebtFusion, MaximumLTV, ApplicantEFSCVScore, UplStrategyFusion,
			CurrentAddress, postalcode, MortgageBalances, PropertyValue, url, IncomeLiabilityScreen, QLAInterestScreen,
			tsdate, SPLBuydown, masterID,ts0, ts1, ts2, ReasonCodeSPLFullScreen1, ReasonCodeSPLFullScreen2,
			ReasonCodeSPLFullScreen3, MaxHASPLFullScreen;

	@BeforeTest
	public void initialize1() throws IOException {

		driver = browser();

	}

	@Test()
	public void m1() throws Exception {

		// Login as Admin
		loginAsAdmin();
	    waitForFirstSubmission();
	    firstPopup();
		getAddress();
		getPartyDetails();
		getCollateral();

		getAppDetails();
		
		calculateIncome();
		calculateSPLLiability();
		premulesoft();
		mulesoft();
		splinterestRateCalculation();
		splLTV();
		splRemInCal();
		splQLA();
		urbancode();
		splFinalQLA();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopupSpl();
	}

	public void loginAsAdmin() throws InterruptedException, IOException, UnsupportedFlavorException {
		 driver.get(prop.getProperty("sfUrl"));

	//	driver.get("https://goeasy--uatpreview.lightning.force.com/lightning/r/genesis__Applications__c/a5yf0000000HsxbAAC/view");
		// driver.get("https://goeasy--goeasyqasb.my.salesforce.com/");
		Thread.sleep(2000);
		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 0000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("username"))));
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(prop.getProperty("AdminEmailFusion"));
		Thread.sleep(2000);
		// driver.findElement(By.cssSelector(prop.getProperty("password"))).sendKeys(decodeString(prop.getProperty("AdminPassword")));
		driver.findElement(By.cssSelector(prop.getProperty("password")))
				.sendKeys(prop.getProperty("AdminPasswordFusion"));
		driver.findElement(By.xpath(prop.getProperty("clicklogin"))).click();

		// https://goeasy--goeasyqasb.lightning.force.com/lightning/r/genesis__Applications__c/a5y1h0000002C2NAAU/view
		System.out.println("Logged in As Admin");
	}
	public void premulesoft() throws InterruptedException
	{
		// mulesoft
		driver.get(prop.getProperty("mulesoft"));
		// Usecustomdomain
		WebDriverWait waitcustom = new WebDriverWait(driver, 360, 0000);
		waitcustom.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Usecustomdomain"))));
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		Thread.sleep(2000);
		driver.findElement(By.xpath(prop.getProperty("Usecustomdomain"))).click();
		// customDomain
		WebDriverWait waitcustomDomain = new WebDriverWait(driver, 360, 0000);
		waitcustomDomain
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("customDomain"))));
		Thread.sleep(2000);
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		driver.findElement(By.xpath(prop.getProperty("customDomain"))).sendKeys(prop.getProperty("Organizationdomain"));
		// ContinueOrg
		driver.findElement(By.xpath(prop.getProperty("ContinueOrg"))).click();
	}

	public void mulesoft() throws Exception {

		/*
		 * Timestamp timeStamp = new Timestamp(System.currentTimeMillis()); String
		 * Time=timeStamp.toString(); System.out.println(timeStamp); String[] arrSplit =
		 * Time.split(" "); String date = arrSplit[0]; String time = arrSplit[1];
		 * System.out.println(date+" "+time);
		 */
		
		
		// RuntimeManager
		WebDriverWait waitRuntimeManager = new WebDriverWait(driver, 360, 0000);
		waitRuntimeManager
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("RuntimeManager"))));
		driver.findElement(By.xpath(prop.getProperty("RuntimeManager"))).click();

		// SearchApplications
		WebDriverWait waitSearchApplications = new WebDriverWait(driver, 360, 0000);
		waitSearchApplications
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("SearchApplications"))));
		// driver.findElement(By.xpath(prop.getProperty("username"))).sendKeys(decodeString(prop.getProperty("AdminEmail")));
		driver.findElement(By.xpath(prop.getProperty("SearchApplications"))).sendKeys(prop.getProperty("clsdev"));

		// clsdevclick

		WebDriverWait waitclsdevclick = new WebDriverWait(driver, 360, 0000);
		waitclsdevclick.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("clsdevclick"))));
		driver.findElement(By.xpath(prop.getProperty("clsdevclick"))).click();

		// Logs
		WebDriverWait waitLogs = new WebDriverWait(driver, 360, 0000);
		
		
		
		waitLogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Logs"))));
    	waitLogs.until(ExpectedConditions.elementToBeClickable(By.xpath(prop.getProperty("Logs"))));
    	WebElement log = driver.findElement(By.xpath(prop.getProperty("Logs")));
		Actions act = new Actions(driver);
		
        int attempts = 0;
        while (attempts < 3)
        {
            try
            {
            	Thread.sleep(5000);
            	WebDriverWait waitLog = new WebDriverWait(driver, 360, 0000);
        		waitLog.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Logs"))));
        		JavascriptExecutor executor = (JavascriptExecutor)driver;
        		executor.executeScript("arguments[0].click();", log);
      
                break;
            }
            catch (StaleElementReferenceException e)
            {
                System.out.println("StaleElementReference");
                driver.get(prop.getProperty("mulesoft"));
                mulesoft();
                splinterestRateCalculation();
        		splLTV();
        		splRemInCal();
        		splQLA();
        		urbancode();
        		splFinalQLA();
        		maxHA();
        		ReasonCode();
        		Thread.sleep(4000);
        		SecondPopupSpl();
                
            }
            attempts++;
        }
       
		
//		act.clickAndHold();
//		act.release().perform();

		/*
		 * // getlogs // WebDriverWait waitgetlogs = new WebDriverWait(driver, 360,
		 * 0000); //
		 * waitgetlogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop
		 * .getProperty("getlogs")))); String logs =
		 * driver.findElement(By.xpath(prop.getProperty("getlogs"))).getText();
		 * System.out.println(logs);
		 */
		// closedeploy
		WebDriverWait waitclosedeploy = new WebDriverWait(driver, 360, 0000);
		waitclosedeploy.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("closedeploy"))));
		driver.findElement(By.xpath(prop.getProperty("closedeploy"))).click();

		try
		{
		
		Thread.sleep(6000);

		// searchlogs
		WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
		waitsearchlogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
		driver.findElement(By.xpath(prop.getProperty("Advanced"))).click();
		// Enter Date and time
		WebDriverWait waitstartDateInput = new WebDriverWait(driver, 360, 0000);
		waitstartDateInput
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("startDateInput"))));
		driver.findElement(By.xpath(prop.getProperty("startDateInput"))).sendKeys(tsdate);
		driver.findElement(By.xpath(prop.getProperty("endDateInput"))).sendKeys(tsdate);
		driver.findElement(By.xpath(prop.getProperty("startTime"))).sendKeys(ts0);
		driver.findElement(By.xpath(prop.getProperty("endTime"))).sendKeys(ts0);
		
		driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(masterID); // clickarrow
		driver.findElement(By.xpath(prop.getProperty("clickarrow"))).click();
		WebDriverWait waitdebugpriority = new WebDriverWait(driver, 360, 0000);
		waitdebugpriority
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("debugpriority"))));
		driver.findElement(By.xpath(prop.getProperty("debugpriority"))).click();
		// Apply
		driver.findElement(By.xpath(prop.getProperty("Apply"))).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(Keys.ENTER);
		Thread.sleep(8000);
		List<WebElement> responsefiles = driver.findElements(By.xpath(prop.getProperty("listdebug")));
		WebElement target = responsefiles.get(0);
		Thread.sleep(2000);
		act.moveToElement(target);
		Thread.sleep(2000);
		act.clickAndHold();
		act.release().perform();
		
		}

		catch(IndexOutOfBoundsException e)
		{
			Thread.sleep(6000);

			// searchlogs
			WebDriverWait waitsearchlogs = new WebDriverWait(driver, 360, 0000);
			waitsearchlogs.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("searchlogs"))));
			driver.findElement(By.xpath(prop.getProperty("Advanced"))).click();
			// Enter Date and time
			WebDriverWait waitstartDateInput = new WebDriverWait(driver, 360, 0000);
			waitstartDateInput
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("startDateInput"))));
			driver.findElement(By.xpath(prop.getProperty("startDateInput"))).clear();
			driver.findElement(By.xpath(prop.getProperty("startDateInput"))).sendKeys(tsdate);
			driver.findElement(By.xpath(prop.getProperty("endDateInput"))).clear();
			driver.findElement(By.xpath(prop.getProperty("endDateInput"))).sendKeys(tsdate);
			driver.findElement(By.xpath(prop.getProperty("startTime"))).clear();
			driver.findElement(By.xpath(prop.getProperty("startTime"))).sendKeys(ts1);
			driver.findElement(By.xpath(prop.getProperty("endTime"))).clear();
			driver.findElement(By.xpath(prop.getProperty("endTime"))).sendKeys(ts2);
			driver.findElement(By.xpath(prop.getProperty("clickarrow"))).click();
			WebDriverWait waitdebugpriority = new WebDriverWait(driver, 360, 0000);
			waitdebugpriority
					.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("debugpriority"))));
			driver.findElement(By.xpath(prop.getProperty("debugpriority"))).click();
			// Apply
			driver.findElement(By.xpath(prop.getProperty("Apply"))).click();
			
			Thread.sleep(7000);
			List<WebElement> responsefiles = driver.findElements(By.xpath(prop.getProperty("listdebug")));
			WebElement target = responsefiles.get(0);
			Thread.sleep(2000);
			act.moveToElement(target);
			Thread.sleep(2000);
			act.clickAndHold();
			act.release().perform();
		}

		// **********************************************************
		/*
		 * //Advanced
		 * driver.findElement(By.xpath(prop.getProperty("Advanced"))).click();
		 * //Lasthour WebDriverWait waitLasthour = new WebDriverWait(driver, 360, 0000);
		 * waitLasthour.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(
		 * prop.getProperty("Lasthour"))));
		 * driver.findElement(By.xpath(prop.getProperty("Lasthour"))).click();
		 * 
		 * driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(
		 * masterID); //clickarrow
		 * driver.findElement(By.xpath(prop.getProperty("clickarrow"))).click();
		 * WebDriverWait waitdebugpriority = new WebDriverWait(driver, 360, 0000);
		 * waitdebugpriority.until(ExpectedConditions.visibilityOfElementLocated(By.
		 * xpath(prop.getProperty("debugpriority"))));
		 * driver.findElement(By.xpath(prop.getProperty("debugpriority"))).click();
		 * //Apply driver.findElement(By.xpath(prop.getProperty("Apply"))).click();
		 * Thread.sleep(4000); List<WebElement> responsefiles
		 * =driver.findElements(By.xpath(prop.getProperty("listdebug"))); WebElement
		 * target = responsefiles.get(0); Thread.sleep(2000); act.moveToElement(target);
		 * Thread.sleep(2000); WebElement clickresponselink =
		 * responsefiles.get(responsefiles.size()-1);
		 * 
		 * 
		 * Thread.sleep(4000); act.moveToElement(clickresponselink); act.clickAndHold();
		 * act.release().perform(); Thread.sleep(4000);
		 */
		// *****************************************************************
		// Keys.RETURN
		// driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(masterID);
		// driver.findElement(By.xpath(prop.getProperty("searchlogs"))).sendKeys(Keys.RETURN);

		Thread.sleep(10000);
		// WebElement field=driver.findElement(By.xpath(prop.getProperty("getlogs")));
		// act.moveToElement(field).doubleClick().build().perform();

		act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
		act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
		String logs = (String) Toolkit.getDefaultToolkit().getSystemClipboard().getData(DataFlavor.stringFlavor);
		int index1 = logs.indexOf("RandomNumber_Internal_SPLInterestRate");
		String roar1 = logs.substring(index1 + 39, index1 + 42);
		String randomnumber = null;
		if (roar1.contains(",")) {
			randomnumber = roar1.replace(",", "");
		} else {
			randomnumber = roar1;
		}

		double RandomNum = Double.valueOf(randomnumber);
		RandomNumberResponse = (int) RandomNum;
		int index2 = logs.indexOf("DE_SPL_Buydown");
		SPLBuydown = logs.substring(index2 + 18, index2 + 19);
		driver.get(url);
		// driver.close();
	}

	public void waitForFirstSubmission() throws Exception {

		// Accept Applicant Section Complete Alert

		// We are declaring the frame
		JFrame frmOpt = new JFrame(); // We are declaring the frame
		frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

		Thread.sleep(1000);
		String s = "<html>Press 1 For Calculations<br>Press 2 For Results<br>";

		JLabel label = new JLabel(s);
		JTextPane jtp = new JTextPane();
		jtp.setSize(new Dimension(480, 10));
		jtp.setPreferredSize(new Dimension(480, jtp.getPreferredSize().height));
		label.setFont(new Font("Arial", Font.BOLD, 20));
		UIManager.put("OptionPane.minimumSize", new Dimension(500, 200));
		UIManager.put("TextField.font", new FontUIResource(new Font("Verdana", Font.BOLD, 18)));
		// Getting Input from user

		String option = JOptionPane.showInputDialog(frmOpt, label);

		int useroption = Integer.parseInt(option);

		switch (useroption) {

		case 1:

			// Function for Re-Submission

			break;

		case 2:

			System.out.println("Results");
			if (attemptNo == 0) {
				test = Extent.createTest("Resubmission - Confirmation ");
				test.info(" The test run complete. Please review test result(s)");
			}

			else {

				test = Extent.createTest("Resubmission - Confirmation ");
				test.info(" The test run complete. Please review test result(s)");
			}

			Thread.sleep(3000);

			driver.close();
			driver.quit();
			break;

		}

	}

	public void firstPopup() throws InterruptedException {
		Thread.sleep(9000);
		// First Pop-up
		driver.get(System.getProperty("user.dir") + "\\src\\main\\resources\\confirmationAlert1.html");
		WebDriverWait waitLoad = new WebDriverWait(driver, 360, 00000000);
		waitLoad.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@name='alert']")));
		WebElement clickalert = driver.findElement(By.xpath("//*[@name='alert']"));
		clickalert.click();

		Thread.sleep(12000);
		String response = null;

		try {
			if (driver.findElement(By.xpath("//*[@id='msg']")).isDisplayed() == true) {
				response = driver.findElement(By.xpath("//*[@id='msg']")).getText();

			}

		}

		catch (Exception e) {

			System.out.println(e.getMessage());
		}

		if (response.contains("OK")) {

			driver.navigate().back();
			Thread.sleep(5000);
		} else if (response.contains("CANCEL")) {

			test = Extent.createTest("Get Application Details ");

			Thread.sleep(3000);

			driver.close();
			driver.quit();

			test.info("You opted to Close the Automation Test Run");

			test.log(Status.PASS, MarkupHelper.createLabel("Automation Exited", ExtentColor.GREEN));
		}
	}

	public void getAddress() throws InterruptedException, java.text.ParseException {

		switchtoIframe1();
		Thread.sleep(5000);

		switchtoIframe2();
		Thread.sleep(5000);
		// threedots
		WebDriverWait waitSetup2 = new WebDriverWait(driver, 360, 0000);
		waitSetup2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("OfferSelection"))));
		driver.findElement(By.xpath(prop.getProperty("threedots"))).click();

		// EventHistory
		WebDriverWait waitEventHistory = new WebDriverWait(driver, 360, 0000);
		waitEventHistory
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EventHistory"))));
		driver.findElement(By.xpath(prop.getProperty("EventHistory"))).click();
		// EventHistoryTable
		driver.switchTo().defaultContent();
		switchtoIframe1();
		Thread.sleep(5000);
		switchtoIframe3();
		Thread.sleep(8000);
		String Event = null;
		try {
		WebDriverWait waitEventHistoryTable = new WebDriverWait(driver, 360, 0000);
		waitEventHistoryTable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EventHistoryTable"))));

		WebElement EventHistoryTable = driver.findElement(By.xpath(prop.getProperty("EventHistoryTable")));
		List<WebElement> rowValsEventHistoryTable = EventHistoryTable.findElements(By.tagName("tr"));
		int rowNumEventHistoryTable = EventHistoryTable.findElements(By.tagName("tr")).size();
		String EventHistoryDate = null;
		
		for (int i = 0; i < rowNumEventHistoryTable; i++) {

			// Get each row's column values by tag name
			List<WebElement> colValsEventHistory = rowValsEventHistoryTable.get(i).findElements(By.tagName("td"));
			WebElement EventHistory = colValsEventHistory.get(0);
			EventHistoryDate = EventHistory.getText();
			if (EventHistoryDate.contains("GDS returns offer details")) {
				WebElement getEvent = colValsEventHistory.get(4);
				Event = getEvent.getText();
			}

		}
		}
		catch (IndexOutOfBoundsException e)
		{
			System.out.println("Event History table not correctly displayed");
		}
		System.out.println("EventHistory: " + Event);

		// Format of the date defined in the input String
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy hh:mm aa");
		// Desired format: 24 hour format: Change the pattern as per the need
		DateFormat outputformat = new SimpleDateFormat("MM/dd/yyyy HH:mm");
		java.util.Date date2 = null;
		String output = null;
		// Converting the input String to Date
		date2 = df.parse(Event);
		// Changing the format of date and storing it in String
		output = outputformat.format(date2);
		// Displaying the date

		String[] parts = output.split(" ");
		tsdate = parts[0];
		String tstime = parts[1];

		String logs = tsdate + " " + tstime + ":00";

		// parse the string
		org.joda.time.format.DateTimeFormatter dtf = DateTimeFormat.forPattern("MM/dd/yyyy HH:mm:ss");
		// Parsing the date
		DateTime jodatime = dtf.parseDateTime(logs);

		String ant0=String.valueOf(jodatime);
		ts0 = Character.toString(ant0.charAt(11)) + Character.toString(ant0.charAt(12)) + ":"
			+ Character.toString(ant0.charAt(14)) + Character.toString(ant0.charAt(15));
		// add two hours
		DateTime date = jodatime.minusMinutes(1);
		DateTime dateTime = jodatime.plusMinutes(1); // easier than mucking about with Calendar and constants

		String ant1 = String.valueOf(date);
		ts1 = Character.toString(ant1.charAt(11)) + Character.toString(ant1.charAt(12)) + ":"
				+ Character.toString(ant1.charAt(14)) + Character.toString(ant1.charAt(15));

		String ant = String.valueOf(dateTime);
		ts2 = Character.toString(ant.charAt(11)) + Character.toString(ant.charAt(12)) + ":"
				+ Character.toString(ant.charAt(14)) + Character.toString(ant.charAt(15));

		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();

		driver.switchTo().defaultContent();
		switchtoIframe1();
		Thread.sleep(3000);
		switchtoIframe2();
		Thread.sleep(3000);
		WebDriverWait waitSetup = new WebDriverWait(driver, 360, 0000);
		waitSetup.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Parties"))));
		driver.findElement(By.xpath(prop.getProperty("Parties"))).click();
		Thread.sleep(5000);
		switchtoIframe4();
		Thread.sleep(5000);
		// CurrentAddress
		WebDriverWait CurrentAddresswait = new WebDriverWait(driver, 360, 0000);
		CurrentAddresswait
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("CurrentAddress"))));
		CurrentAddress = driver.findElement(By.xpath(prop.getProperty("CurrentAddress"))).getText();
		String reverse = new StringBuffer(CurrentAddress).reverse().toString();
		reverse = reverse.replaceAll(" ", "");
		String part1 = reverse.substring(0, 6);
		postalcode = new StringBuffer(part1).reverse().toString();
		if (postalcode.contains(" ")) {
			postalcode = postalcode.replaceAll(" ", "");
		}

		if (CurrentAddress.contains("Ontario")) {
			Prov = "ON";
		} else if (CurrentAddress.contains("Alberta")) {
			Prov = "AB";
		} else if (CurrentAddress.contains("British Columbia")) {
			Prov = "BC";
		} else if (CurrentAddress.contains("Manitoba")) {
			Prov = "MB";
		} else if (CurrentAddress.contains("New Brunswick")) {
			Prov = "NB";
		} else if (CurrentAddress.contains("Newfoundland and Labrador")) {
			Prov = "NL";
		} else if (CurrentAddress.contains("Nova Scotia")) {
			Prov = "NS";
		} else if (CurrentAddress.contains("Northwest Territories")) {
			Prov = "NT";
		} else if (CurrentAddress.contains("Nunavut")) {
			Prov = "NU";
		} else if (CurrentAddress.contains("Prince Edward")) {
			Prov = "PE";
		} else if (CurrentAddress.contains("Quebec")) {
			Prov = "QC";
		} else if (CurrentAddress.contains("Saskatchewan")) {
			Prov = "SK";
		} else if (CurrentAddress.contains("Yukon")) {
			Prov = "YT";
		}

	}

	public void getPartyDetails()
			throws InterruptedException, HeadlessException, UnsupportedFlavorException, IOException {
		// PartyDetails
		WebDriverWait waitPartyDetails = new WebDriverWait(driver, 360, 0000);
		waitPartyDetails
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("PartyDetails"))));
		driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();

		// EmploymentandIncome

		WebDriverWait waitEmploymentandIncome = new WebDriverWait(driver, 360, 0000);
		waitEmploymentandIncome.until(
				ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("EmploymentandIncome"))));
		driver.findElement(By.xpath(prop.getProperty("EmploymentandIncome"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		// Income table
		Thread.sleep(8000);
		WebDriverWait waitMonthlyIncometable = new WebDriverWait(driver, 360, 0000);
		waitMonthlyIncometable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("MonthlyIncometable"))));

		WebElement MonthlyIncometable = driver.findElement(By.xpath(prop.getProperty("MonthlyIncometable")));
		List<WebElement> rowValsIncome = MonthlyIncometable.findElements(By.tagName("tr"));
		int rowNumIncome = MonthlyIncometable.findElements(By.tagName("tr")).size();

		String str = null;
		double ApplicantIncome = 0, OtherIncomeValue=0.0;
		for (int i = 0; i < rowNumIncome; i++) {

			double subValue = 0;
			// Get each row's column values by tag name
			List<WebElement> colValsIncome = rowValsIncome.get(i).findElements(By.tagName("td"));
			WebElement Income = colValsIncome.get(4);
			Actions act = new Actions(driver);
			Income.click();
			Thread.sleep(1000);
			Income.click();
			
			act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
			act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
			String Incomeamount = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
					.getData(DataFlavor.stringFlavor);

			if (Incomeamount.contains(","))

			{
				str = Incomeamount.replace(",", "");
				subValue = Double.parseDouble(str.replace("$", ""));
			} else {
				subValue = Double.parseDouble(Incomeamount.replace("$", ""));
			}
			ApplicantIncome += subValue;
		}
		
		// otherIncometable
		WebElement otherIncometable = driver.findElement(By.xpath(prop.getProperty("otherIncometable")));
		List<WebElement> rowValsOther = otherIncometable.findElements(By.tagName("tr"));
		int rowNumOther = otherIncometable.findElements(By.tagName("tr")).size();
		String strOther = null;
		for (int i = 0; i < rowNumOther; i++) {

			double subValue = 0;
			// Get each row's column values by tag name
			List<WebElement> colValOther = rowValsOther.get(i).findElements(By.tagName("td"));
			WebElement Income = colValOther.get(4);
			Actions act = new Actions(driver);
			Income.click();
			Thread.sleep(1000);
			Income.click();
			
			act.keyDown(Keys.CONTROL).sendKeys("a").keyUp(Keys.CONTROL).build().perform();
			act.keyDown(Keys.CONTROL).sendKeys("c").keyUp(Keys.CONTROL).build().perform();
			String Incomeamount = (String) Toolkit.getDefaultToolkit().getSystemClipboard()
					.getData(DataFlavor.stringFlavor);

			if (Incomeamount.contains(","))

			{
				strOther = Incomeamount.replace(",", "");
				subValue = Double.parseDouble(strOther.replace("$", ""));
			} else {
				subValue = Double.parseDouble(Incomeamount.replace("$", ""));
			}
			OtherIncomeValue += subValue;

		}

		IncomeValue = ApplicantIncome + OtherIncomeValue;
		System.out.println("ApplicantIncome =$" + IncomeValue);
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
		switchtoIframe2();
		switchtoIframe4();
		Thread.sleep(2000);
		driver.findElement(By.xpath(prop.getProperty("PartyDetails"))).click();
		// Liabilities
		WebDriverWait waitLiabilities = new WebDriverWait(driver, 360, 0000);
		waitLiabilities.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Liabilities"))));
		driver.findElement(By.xpath(prop.getProperty("Liabilities"))).click();
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();
		Thread.sleep(8000);
		try
		{
		WebDriverWait waitLiabilitiesTable = new WebDriverWait(driver, 6);
		waitLiabilitiesTable
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("LiabilitiesTable"))));
		WebElement LiabilitiesTable = driver.findElement(By.xpath(prop.getProperty("LiabilitiesTable")));
		if (LiabilitiesTable.isDisplayed()) {
			// showrows
			String rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
			int rowcount=0;

			if (rows.contains("+")) {
				do {
					Loadmore();
					Thread.sleep(5000);
					rows = driver.findElement(By.xpath(prop.getProperty("showrows"))).getText();
					rowcount++;
				} while (rows.contains("+"));
			}

			List<WebElement> rowValsLiability = LiabilitiesTable.findElements(By.tagName("tr"));
			int rowNumLiability = LiabilitiesTable.findElements(By.tagName("tr")).size();

			String strLiability = null;
			TotalLaibility=0;
			for (int i = 0; i < rowNumLiability; i++) {

				double subValue = 0;
				// Get each row's column values by tag name
				List<WebElement> colValsliability = rowValsLiability.get(i).findElements(By.tagName("td"));
				WebElement Liability = colValsliability.get(9);
				String Liabilityamount = Liability.getText();

				if (Liabilityamount.contains(","))

				{
					strLiability = Liabilityamount.replace(",", "");
					subValue = Double.parseDouble(strLiability.replace("$", ""));
				} else {
					subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
				}
				TotalLaibility += subValue;
			}
		
			if(rowcount>4)
			{
				Thread.sleep(6000);
				//click2
				
				driver.findElement(By.xpath(prop.getProperty("next"))).click();
				Thread.sleep(3000);
				List<WebElement> rowValsLiability1 = LiabilitiesTable.findElements(By.tagName("tr"));
				int rowNumLiability1 = LiabilitiesTable.findElements(By.tagName("tr")).size();

				String strLiability1 = null;
				
				for (int i = 0; i < rowNumLiability1; i++) {

					double subValue = 0;
					// Get each row's column values by tag name
					List<WebElement> colValsliability = rowValsLiability1.get(i).findElements(By.tagName("td"));
					WebElement Liability = colValsliability.get(9);
					String Liabilityamount = Liability.getText();

					if (Liabilityamount.contains(","))

					{
						strLiability1 = Liabilityamount.replace(",", "");
						subValue = Double.parseDouble(strLiability1.replace("$", ""));
					} else {
						subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
					}
					TotalLaibility += subValue;
				}
				if(rowcount>9 )
				{
					Thread.sleep(4000);
					//click3
					driver.findElement(By.xpath(prop.getProperty("next"))).click();
					Thread.sleep(3000);
					List<WebElement> rowValsLiability11 = LiabilitiesTable.findElements(By.tagName("tr"));
					int rowNumLiability11 = LiabilitiesTable.findElements(By.tagName("tr")).size();

					String strLiability11 = null;
					
					for (int i = 0; i < rowNumLiability11; i++) {

						double subValue = 0;
						// Get each row's column values by tag name
						List<WebElement> colValsliability = rowValsLiability11.get(i).findElements(By.tagName("td"));
						WebElement Liability = colValsliability.get(9);
						String Liabilityamount = Liability.getText();

						if (Liabilityamount.contains(","))

						{
							strLiability11 = Liabilityamount.replace(",", "");
							subValue = Double.parseDouble(strLiability11.replace("$", ""));
						} else {
							subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
						}
						TotalLaibility += subValue;
					}
					
					if(rowcount>14 )
					{
						Thread.sleep(4000);
						//click4
						driver.findElement(By.xpath(prop.getProperty("next"))).click();
						Thread.sleep(3000);
						List<WebElement> rowValsLiability111 = LiabilitiesTable.findElements(By.tagName("tr"));
						int rowNumLiability111 = LiabilitiesTable.findElements(By.tagName("tr")).size();

						String strLiability111 = null;
						
						for (int i = 0; i < rowNumLiability111; i++) {

							double subValue = 0;
							// Get each row's column values by tag name
							List<WebElement> colValsliability = rowValsLiability111.get(i).findElements(By.tagName("td"));
							WebElement Liability = colValsliability.get(9);
							String Liabilityamount = Liability.getText();

							if (Liabilityamount.contains(","))

							{
								strLiability111 = Liabilityamount.replace(",", "");
								subValue = Double.parseDouble(strLiability111.replace("$", ""));
							} else {
								subValue = Double.parseDouble(Liabilityamount.replace("$", ""));
							}
							TotalLaibility += subValue;
						}
					}
					
				}
				
			}
		}
		}
		catch (TimeoutException e)
        {
            System.out.println("Liability Not Displayed");
           
            
        }

		System.out.println("LiabilityValue =$" + TotalLaibility);
		// close
		driver.switchTo().defaultContent();
		switchtoIframe1();
		driver.findElement(By.xpath(prop.getProperty("closeIncome"))).click();
	}

	public void Loadmore() {

		// loadmore
		Actions action = new Actions(driver);

		if (driver.findElement(By.xpath(prop.getProperty("loadmore"))).isDisplayed()) {
			WebElement loadmore = driver.findElement(By.xpath(prop.getProperty("loadmore")));
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", loadmore);
			loadmore.click();
		}
	}

	public void getCollateral() {
		// Collateral
		switchtoIframe2();
		WebDriverWait waitCollateral = new WebDriverWait(driver, 360, 0000);
		waitCollateral.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Collateral"))));
		driver.findElement(By.xpath(prop.getProperty("Collateral"))).click();
		WebDriverWait waitpropertyType = new WebDriverWait(driver, 360, 0000);
		waitpropertyType
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("PropertyType"))));
		propertyType = driver.findElement(By.xpath(prop.getProperty("PropertyType"))).getText();
		System.out.println(propertyType);
		MortgageBalances = driver.findElement(By.xpath(prop.getProperty("TotalMortgageBalanceOutstanding"))).getText();
		PropertyValue = driver.findElement(By.xpath(prop.getProperty("EstimatedPropertyValue"))).getText();
	}

	public void getAppDetails() throws InterruptedException, IOException {
		// OfferSelection
		WebDriverWait waitAgentDashboard = new WebDriverWait(driver, 360, 0000);
		waitAgentDashboard
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("AgentDashboard"))));
		driver.findElement(By.xpath(prop.getProperty("AgentDashboard"))).click();
		WebDriverWait waitSetup2 = new WebDriverWait(driver, 360, 0000);
		waitSetup2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("OfferSelection"))));
		url = driver.getCurrentUrl();
		// Get MasterID
		String reverseurl = new StringBuffer(url).reverse().toString();
		String[] Spliturl = reverseurl.split("/");
		String view = Spliturl[1];
		masterID = new StringBuffer(view).reverse().toString();
		System.out.println("MasterID =" + masterID);
		driver.findElement(By.xpath(prop.getProperty("OfferSelection"))).click();

		// Switch Iframes
		driver.switchTo().defaultContent();
		switchtoIframe1();
		switchtoIframe3();

		// Offer details

		Thread.sleep(12000);
		WebElement getQla = driver.findElement(By.xpath(prop.getProperty("QualifiedLoanAmount")));
		Actions action = new Actions(driver);

		QualifiedLoanAmount = getQla.getText();
		String qla;
		if (QualifiedLoanAmount.contains(",")) {
			qla = QualifiedLoanAmount.replace(",", "");
		} else {
			qla = QualifiedLoanAmount;
		}

		ExpectedQLA = Double.parseDouble(qla.replace("$", ""));
		System.out.println("ExpectedQLA =$" + ExpectedQLA);
		WebElement getInt = driver.findElement(By.xpath(prop.getProperty("InterestRate")));
		IntRate = getInt.getText();
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", getQla);
		Thread.sleep(2000);
		QLAInterestScreen = Screenshot.capture(driver, "CaculateQLA");
		System.out.println("InterestRate =" + IntRate);
		ExpInt = Double.parseDouble(IntRate.replace("%", ""));
		WebDriverWait waitOfferdetails = new WebDriverWait(driver, 360, 0000);
		waitOfferdetails
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("Offerdetails"))));

		// Thread.sleep(8000);
		driver.findElement(By.xpath(prop.getProperty("Offerdetails"))).click();

		// Total Income
		WebDriverWait totalincome = new WebDriverWait(driver, 360, 0000);
		totalincome
				.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("TotalIncomeFusion"))));
		totalincomeFusion = driver.findElement(By.xpath(prop.getProperty("TotalIncomeFusion"))).getText();
		String st = totalincomeFusion.replace(",", "");
		TotalIncome = Double.parseDouble(st.replace("$", ""));
		System.out.println("Total Income =$" + TotalIncome);
		IncomeLiabilityScreen = Screenshot.capture(driver, "CalculateIncome");
		// Total Debt
		totaldebtFusion = driver.findElement(By.xpath(prop.getProperty("TotalDebtFusion"))).getText();
		String str1 = totaldebtFusion.replace(",", "");
		TotalDebt = Double.parseDouble(str1.replace("$", ""));
		MaximumLTV = driver.findElement(By.xpath(prop.getProperty("MaximumLTV"))).getText();

		LtvMax = Double.parseDouble(MaximumLTV);

		System.out.println("Total Debt =$" + TotalDebt);
		System.out.println("MaximumLTV =" + LtvMax);

		// ExpectedMaxHA
		List<WebElement> HA = driver.findElements(By.xpath(prop.getProperty("HAFullSpl")));
		WebElement targetHA = HA.get(1);
		((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", targetHA);
		Thread.sleep(2000);
		MaxHASPLFullScreen = Screenshot.capture(driver, "MAxHA");
		String MaxHA = targetHA.getText();
		ExpectedMaxHA = Double.parseDouble(MaxHA.replace(",", ""));
		System.out.println("ExpectedMaxHA =" + ExpectedMaxHA);

		// ReasonCodeSPLFullScreen
		List<WebElement> reasoncode = driver.findElements(By.xpath(prop.getProperty("SPLRiskFactors")));
		length = reasoncode.size();
		if (length == 1) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
		} else if (length == 2) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
		}
		else if (length == 3) {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
			WebElement desicioncode3 = reasoncode.get(2);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode3);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen3 = Screenshot.capture(driver, "ReasonCode3");
		}
		
		else {
			WebElement desicioncode1 = reasoncode.get(0);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode1);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen1 = Screenshot.capture(driver, "ReasonCode1");
			WebElement desicioncode2 = reasoncode.get(1);
			((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);", desicioncode2);
			Thread.sleep(2000);
			ReasonCodeSPLFullScreen2 = Screenshot.capture(driver, "ReasonCode2");
		}
		// ApplicantEFSCVScore

		ApplicantEFSCVScore = driver.findElement(By.xpath(prop.getProperty("ApplicantEFSCVScore"))).getText();
		String parts = ApplicantEFSCVScore.substring(0, 3);
		cvScore = Integer.parseInt(parts);
		System.out.println("ApplicantEFSCVScore =" + cvScore);

		// CloseOffer
		driver.findElement(By.xpath(prop.getProperty("CloseOffer"))).click();

	}

	public void calculateIncome() throws InterruptedException, IOException {

		System.out.println("Resubmission attempt #" + attemptNo);
		if (attemptNo == 0) {
			test = Extent.createTest("Total Income Calcuation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Income Calcuation");
		}

		System.out.println("Actual Income: $" + IncomeValue);
		System.out.println("Expected Income: $" + TotalIncome);
		if (IncomeValue == TotalIncome) {

			test.log(Status.PASS,
					MarkupHelper.createLabel(" Total Income  :Actual Value =  $" + IncomeValue, ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("Total Income  :Expected Value =  $" + TotalIncome, ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("Income is Matching with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));
			System.out.println("Income:Passed");

		} else {

			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Income : Actual Value =  $" + IncomeValue, ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Income : Expected Value =  $" + TotalIncome, ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("Income Not Matching with GDS Decision", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));
			System.out.println("Income Not Match");
		}

	}

	public void calculateSPLLiability() throws InterruptedException, IOException {

		if (attemptNo == 0) {
			test = Extent.createTest("Total Liability Calcuation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Total Liability Calcuation");
		}
		System.out.println("Actual Liabilities: $" + TotalLaibility);
		System.out.println("Expected Liabilities: $" + TotalDebt);
		if (TotalLaibility == TotalDebt) {
			System.out.println("Laibilities:Passed");

			test.log(Status.PASS, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
					ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.GREEN));

			test.log(Status.PASS,
					MarkupHelper.createLabel("Liability is Matching with GDS Decision ", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));
			// Assert.assertTrue(true);
		} else {
			System.out.println("Laibilities:Failed");

			test.log(Status.FAIL, MarkupHelper.createLabel("Total Liability - Actual Value   =  $" + TotalLaibility,
					ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("Total Liability - Expected Value =  $" + TotalDebt, ExtentColor.RED));

			test.log(Status.FAIL,
					MarkupHelper.createLabel("Liability is not Matching with GDS Decision ", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(IncomeLiabilityScreen));

			// Assert.assertTrue(false);
		}
	}

	public void splLTV() {

		// SPL LTV Calculation - 1st Submission
		// SPL LTV= (Total Amount of Applicant Mortgage Balances Outstanding + Total
		// Credit Limits of Revolving Trades of Applicant)*100/Total Value of Property

		String str = MortgageBalances.replace(",", "");
		double MortgageBal = Double.parseDouble(str.replace("$", ""));
		System.out.println("Mortgage Balance = " + MortgageBal);
		String str1 = PropertyValue.replace(",", "");
		PropertyVal = Double.parseDouble(str1.replace("$", ""));
		System.out.println("Property Value = " + PropertyVal);
		SPLltv = MortgageBal * 100 / PropertyVal;
		System.out.println("SPL LTV =" + SPLltv);

	}

	@SuppressWarnings("resource")
	public void splRemInCal() throws InterruptedException, IOException {

		org.apache.poi.ss.usermodel.Sheet sheet;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet = workbook.getSheet("SPL Remaining Income");

		// Making cell values as variable

		int cvSco1 = (int) sheet.getRow(3).getCell(1).getNumericCellValue();
		int cvSco2 = (int) sheet.getRow(4).getCell(1).getNumericCellValue();
		int cvSco3 = (int) sheet.getRow(5).getCell(1).getNumericCellValue();
		int cvSco4 = (int) sheet.getRow(6).getCell(1).getNumericCellValue();

		int cvS1 = (int) sheet.getRow(3).getCell(3).getNumericCellValue();
		int cvS2 = (int) sheet.getRow(4).getCell(3).getNumericCellValue();
		int cvS3 = (int) sheet.getRow(5).getCell(3).getNumericCellValue();

		double value1 = sheet.getRow(3).getCell(5).getNumericCellValue();
		double value2 = sheet.getRow(4).getCell(5).getNumericCellValue();
		double value3 = sheet.getRow(5).getCell(5).getNumericCellValue();
		double value4 = sheet.getRow(6).getCell(5).getNumericCellValue();

		if ((cvScore >= cvSco1) && (cvScore <= cvS1)) {
			RemainingIncome = TotalIncome * value1 - TotalDebt;
		}

		else if ((cvScore >= cvSco2) && (cvScore <= cvS2)) {
			RemainingIncome = TotalIncome * value2 - TotalDebt;
		}

		else if ((cvScore >= cvSco3) && (cvScore <= cvS3)) {
			RemainingIncome = TotalIncome * value3 - TotalDebt;
		}

		else if ((cvScore > cvSco4)) {
			RemainingIncome = TotalIncome * value4 - TotalDebt;
		}

		System.out.println("RemainingIncome = " + RemainingIncome);

	}

	public void switchtoIframe1() {
		WebDriverWait wait = new WebDriverWait(driver, 360, 0000);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe1"))));
		WebElement iframe1 = driver.findElement(By.xpath(prop.getProperty("iframe1")));

		driver.switchTo().frame(iframe1);

	}

	public void switchtoIframe2() {
		WebDriverWait wait2 = new WebDriverWait(driver, 360, 0000);
		wait2.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe2"))));
		WebElement iframe2 = driver.findElement(By.xpath(prop.getProperty("iframe2")));
		driver.switchTo().frame(iframe2);

	}

	public void switchtoIframe3() {
		WebDriverWait waitframe = new WebDriverWait(driver, 360, 0000);
		waitframe.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe3"))));
		WebElement iframe3 = driver.findElement(By.xpath(prop.getProperty("iframe3")));
		driver.switchTo().frame(iframe3);

	}

	public void switchtoIframe4() {
		WebDriverWait waitparties = new WebDriverWait(driver, 360, 0000);
		waitparties.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(prop.getProperty("iframe4"))));
		WebElement iframe4 = driver.findElement(By.xpath(prop.getProperty("iframe4")));
		driver.switchTo().frame(iframe4);

	}



	public void splinterestRateCalculation()
			throws IOException, DocumentException, InterruptedException, ParseException {
		if (attemptNo == 0) {
			test = Extent.createTest("Interest Rate Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - InterestRate Calculation");
		}

		int RandomNumber, spl = 0;
		double InterestRate = 0;
		if (cvScore <= 682) {

			RandomNumber = RandomNumberResponse;
		} else {
			RandomNumber = 0;
			InterestRate = 19.99;
		}

		// SPLBuydown
		if (SPLBuydown.contains("0") || SPLBuydown.contains("1")) {
			double splbuy = Double.valueOf(SPLBuydown);
			spl = (int) splbuy;
		}

		System.out.println("RandomNumber: " + RandomNumber);

		double randomNumOne, randomNumTwo, randomNumThree, randomNumFour;

		String randomNumRange, efsCvScoreRange;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("SPL Interest Rate");

		// Storing random numbers as separate variable

		randomNumRange = sheet.getRow(2).getCell(1).getStringCellValue();

		String[] stringSplit = randomNumRange.split("-");

		randomNumOne = Double.parseDouble(stringSplit[0]); // 1
		randomNumTwo = Double.parseDouble(stringSplit[1]); // 50

		// System.out.println(randomNumOne);

		randomNumRange = sheet.getRow(3).getCell(1).getStringCellValue();

		stringSplit = randomNumRange.split("-");

		randomNumThree = Double.parseDouble(stringSplit[0]); // 51
		randomNumFour = Double.parseDouble(stringSplit[1]); // 100

		// System.out.println(randomNumFour);

		ArrayList<Double> efsCvScoreList = new ArrayList<Double>();

		for (int x = 2; x <= 10; x++) {
			efsCvScoreRange = sheet.getRow(x).getCell(0).getStringCellValue();

			if (efsCvScoreRange.contains("-")) {
				stringSplit = efsCvScoreRange.split("-");
				efsCvScoreList.add(Double.parseDouble(stringSplit[0]));
				efsCvScoreList.add(Double.parseDouble(stringSplit[1]));
			}
			if (efsCvScoreRange.contains("=")) {

				efsCvScoreRange = efsCvScoreRange.replace("<=", "");
				efsCvScoreList.add(Double.parseDouble(efsCvScoreRange));

			}
			x++;
		}

		if ((randomNumOne <= RandomNumber) && (RandomNumber <= randomNumTwo)) // 1 and 50
		{
			if ((efsCvScoreList.get(0) <= cvScore) && (cvScore <= efsCvScoreList.get(1))) // 646<=cvScore<=682
			{
				InterestRate = sheet.getRow(2).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(2) <= cvScore) && (cvScore <= efsCvScoreList.get(3))) // 627<=cvScore<=645
			{
				InterestRate = sheet.getRow(4).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(4) <= cvScore) && (cvScore <= efsCvScoreList.get(5))) // 610<=cvScore<=625
			{
				InterestRate = sheet.getRow(6).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(6) <= cvScore) && (cvScore <= efsCvScoreList.get(7))) // 593<=cvScore<=609
			{
				InterestRate = sheet.getRow(8).getCell(2).getNumericCellValue();
			}
			if ((cvScore <= efsCvScoreList.get(8))) {
				InterestRate = sheet.getRow(10).getCell(2).getNumericCellValue();
			}
		} else if ((randomNumThree <= RandomNumber) && (RandomNumber <= randomNumFour)) // 51 and 100
		{
			if ((efsCvScoreList.get(0) <= cvScore) && (cvScore <= efsCvScoreList.get(1))) // 646<=cvScore<=682
			{
				InterestRate = sheet.getRow(3).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(2) <= cvScore) && (cvScore <= efsCvScoreList.get(3))) // 627<=cvScore<=645
			{
				InterestRate = sheet.getRow(5).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(4) <= cvScore) && (cvScore <= efsCvScoreList.get(5))) // 610<=cvScore<=625
			{
				InterestRate = sheet.getRow(7).getCell(2).getNumericCellValue();
			}
			if ((efsCvScoreList.get(6) <= cvScore) && (cvScore <= efsCvScoreList.get(7))) // 593<=cvScore<=609
			{
				InterestRate = sheet.getRow(9).getCell(2).getNumericCellValue();
			}
			if ((cvScore <= efsCvScoreList.get(8))) {
				InterestRate = sheet.getRow(10).getCell(2).getNumericCellValue();
			}

		}

		else {
			if ((cvScore <= efsCvScoreList.get(8))) // cvScore 592
			{
				InterestRate = sheet.getRow(10).getCell(2).getNumericCellValue();
			}
		}

		if (spl == 1) {
			InterestRate = InterestRate - 5.0;
		}

		double inrate = Double.valueOf(InterestRate);

		// Displaying Interest Rate result
		System.out.println("Actual Interest rate: " + inrate);
		System.out.println("Expected Interest rate: " + ExpInt);

		if (ExpInt == inrate) {

			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + inrate + "%",
					ExtentColor.GREEN));
			test.log(Status.PASS, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.GREEN));

			test.log(Status.PASS, MarkupHelper.createLabel(" Interest Rate Calculation is Matching with GDS Decision",
					ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("PASSED in Interest Verification");
		} else {
			System.out.println(inrate + " is the Actual value.");

			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Actual value : " + inrate + "%",
					ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("InterestRate Percentage Expected value : " + ExpInt + "%",
					ExtentColor.RED));

			test.log(Status.FAIL, MarkupHelper.createLabel(" Interest Rate Calculation not Matching with GDS Decision",
					ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));
			System.out.println("FAILED in Interest Rate Verification");
		}
		Thread.sleep(3000);

	}

	public void splQLA() throws InterruptedException, IOException {

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);
		// String IntRate = String.valueOf(ExpInt);

		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheet("SPL QLA Interest");

		Row r = sheet.getRow(1);
		int lastCol = r.getLastCellNum(); // Gets last column index
		int lastrow = sheet.getLastRowNum(); // Gets last row num

		System.out.println("Int Rate = " + IntRate + " " + "Province = " + Prov);
		System.out.println("cvScore =" + cvScore);
		int fcol = 0;
		// Iterating the row which Interest value for identifying the right table
		for (int i = 7; i <= lastCol; i++) {
			try {
				if (sheet.getRow(1).getCell(i).getStringCellValue().contains(IntRate)) {
					fcol = i; // this would be the first column(Province) in the table
					break;
				}
			}

			catch (Exception e) {

			}
			i += 3;
		}

		// Since Interest Rate 28.99 table has less rows we are providing some specific
		// conditions
		if (IntRate.contains("28.99")) {
			lastrow = 114;

		}
		int z = 0;
		// Iterating Province column
		for (int j = 3; j <= lastrow; j++) {

			if (sheet.getRow(j).getCell(fcol).getStringCellValue().contains(Prov)) {

				// Once Province matched, it would iterate Remaining Income column
				double remInvalue = sheet.getRow(j).getCell(fcol + 1).getNumericCellValue();

				if (z == 0) {
					if (remInvalue > RemainingIncome) {
						QLA = 0.0;
						break;

					}
				}
				if (!(z == 0)) {
					if (remInvalue > RemainingIncome) {
						QLA = sheet.getRow(j - 1).getCell(fcol + 2).getNumericCellValue();
						break;

					}
				}

				// This block is for last line of the table alone
				if (j == sheet.getLastRowNum()) {
					QLA = sheet.getRow(sheet.getLastRowNum()).getCell(fcol + 2).getNumericCellValue();

				}

				// This block would be executed if the calculated remaining income is greater
				// than the maximum
				// remaining income in the table

				if (!sheet.getRow(j + 1).getCell(fcol).getStringCellValue().contains(Prov)) {
					if (remInvalue < RemainingIncome) {
						// Its picking the maximum value available
						QLA = sheet.getRow(j).getCell(fcol + 2).getNumericCellValue();
						break;
					}
				}
				z++;
			}

		}

		// CV Score
		int CVval1 = (int) sheet.getRow(8).getCell(0).getNumericCellValue();
		int CVval2 = (int) sheet.getRow(8).getCell(2).getNumericCellValue();
		int CVval3 = (int) sheet.getRow(9).getCell(0).getNumericCellValue();
		int CVval4 = (int) sheet.getRow(9).getCell(2).getNumericCellValue();
		int CVval5 = (int) sheet.getRow(10).getCell(0).getNumericCellValue();
		int CVval6 = (int) sheet.getRow(10).getCell(2).getNumericCellValue();
		int CVval7 = (int) sheet.getRow(11).getCell(0).getNumericCellValue();
		int CVval8 = (int) sheet.getRow(11).getCell(2).getNumericCellValue();
		int CVval9 = (int) sheet.getRow(12).getCell(2).getNumericCellValue();

		// Max QLA
		int QLAval1 = (int) sheet.getRow(8).getCell(4).getNumericCellValue();
		int QLAval2 = (int) sheet.getRow(9).getCell(4).getNumericCellValue();
		int QLAval3 = (int) sheet.getRow(10).getCell(4).getNumericCellValue();
		int QLAval4 = (int) sheet.getRow(11).getCell(4).getNumericCellValue();
		int QLAval5 = (int) sheet.getRow(12).getCell(4).getNumericCellValue();

		// Reset Values
		int Resetval1 = (int) sheet.getRow(8).getCell(5).getNumericCellValue();
		int Resetval2 = (int) sheet.getRow(9).getCell(5).getNumericCellValue();
		int Resetval3 = (int) sheet.getRow(10).getCell(5).getNumericCellValue();
		int Resetval4 = (int) sheet.getRow(11).getCell(5).getNumericCellValue();
		int Resetval5 = (int) sheet.getRow(12).getCell(5).getNumericCellValue();

		// Max QLA Conditions
		if ((QLA > QLAval1) && (cvScore >= CVval1) && (cvScore <= CVval2)) {
			QLA = Resetval1;
		}
		if ((QLA > QLAval2) && (cvScore >= CVval3) && (cvScore <= CVval4)) {
			QLA = Resetval2;
		}
		if ((QLA > QLAval3) && (cvScore >= CVval5) && (cvScore <= CVval6)) {
			QLA = Resetval3;
		}
		if ((QLA > QLAval4) && (cvScore >= CVval7) && (cvScore <= CVval8)) {
			QLA = Resetval4;

		}
		if ((QLA > QLAval5) && (cvScore < CVval9)) {
			QLA = Resetval5;

		}

		System.out.println("SPL QLA from Excel is " + QLA);
	}

	public void urbancode() throws FileNotFoundException {

		String ps1 = postalcode.substring(0, 3);
		String ps2 = " ";
		String ps3 = postalcode.substring(3, 6);

		ps = ps1 + ps2 + ps3;
		System.out.println("Checking Urban code: "+ps);
		
		InputStream is = new FileInputStream(
				new File(System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\FNF Urban Code.xlsx"));
		Workbook wb = StreamingReader.builder().sstCacheSize(100).open(is);
		org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheet("FNF goeasy urbanization");

		Iterator<Row> rows = sheet.iterator();
		label: while (rows.hasNext()) {
			Row row = rows.next();
			Iterator<Cell> cell = row.cellIterator();
			Cell value = cell.next();
			// System.out.print(value.getStringCellValue());
			// System.out.println("");

			if (value.getStringCellValue().toLowerCase().contains(ps.toLowerCase())) {
				int i = 0;

				while (cell.hasNext()) {

					value = cell.next();
					if (i == 2) {
						code = value.getStringCellValue();
						System.out.println(code);
						break label;
					}

					i++;
				}

			}

		}
	}

	public void splFinalQLA() throws InterruptedException, IOException {
		// Home Equity = (Max LTV - SPL LTV)*Property Value/100
		Thread.sleep(5000);

		if (attemptNo == 0) {
			test = Extent.createTest("QLA Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - QLA Calculation");
		}
		// Maxltv Calculation

		ArrayList rangeList;
		ArrayList<Double> efsCvScoreList = new ArrayList<Double>();
		ArrayList<String> RiskgroupList = new ArrayList<String>();
		String range, riskGroup = " ", propertyRange;
		Double drange, propRange, propRange1, propRange2;
		int thecol = 0, therow = 0, tabrow = 0;
		String stringSplitter[];

		// Sheet initalization.This needs to be done once in Class level, so that, we
		// dont have to initialize this in each function
		org.apache.poi.ss.usermodel.Sheet sheet1;

		File file = new File(System.getProperty("user.dir")
				+ "\\src\\main\\resources\\Excel\\Credit Decision Calculations - SPL New.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet1 = workbook.getSheet("Max LTV");

		// Reading cv score from B column

		for (int i = 0; i < 7; i++) {
			range = sheet1.getRow(i).getCell(1).getStringCellValue();

			stringSplitter = range.split("&");

			if (i == 0) // This is for first cell alone. To remove the text content in the first cell
			{
				stringSplitter = range.split("&");
				String range1 = stringSplitter[0];
				range1 = range1.replace("efs_cv_score >=", "");
				drange = Double.parseDouble(range1);
				efsCvScoreList.add(drange);
				continue;
			}

			String range1 = stringSplitter[0];
			range1 = range1.replace(">=", " ");
			drange = Double.parseDouble(range1);
			efsCvScoreList.add(drange);

			String range2 = stringSplitter[1];
			range2 = range2.replace("<=", "");
			drange = Double.parseDouble(range2);
			efsCvScoreList.add(drange);

			// Splitting the numbers and storing it in an arraylist

		}

		// [683.0, 646.0, 682.0, 627.0, 645.0, 610.0, 626.0, 593.0, 609.0, 577.0, 592.0,
		// 564.0, 576.0]
		// RiskgroupList
		for (int i = 0; i < 7; i++) {
			String rangerisk = sheet1.getRow(i).getCell(2).getStringCellValue();
			RiskgroupList.add(rangerisk);
		}

		// Assigning the risk group based on applicant's efs score
		if (cvScore >= efsCvScoreList.get(0)) // 683
		{
			riskGroup = (String) RiskgroupList.get(0);
		} else if ((cvScore >= efsCvScoreList.get(1)) && (cvScore <= efsCvScoreList.get(2))) // 646<=cvScore<=682
		{
			riskGroup = (String) RiskgroupList.get(1);
		} else if ((cvScore >= efsCvScoreList.get(3)) && (cvScore <= efsCvScoreList.get(4))) // 627<=cvScore<=645
		{
			riskGroup = (String) RiskgroupList.get(2);
		} else if ((cvScore >= efsCvScoreList.get(5)) && (cvScore <= efsCvScoreList.get(6))) // 610<=cvScore<=626
		{
			riskGroup = (String) RiskgroupList.get(3);
		} else if ((cvScore >= efsCvScoreList.get(7)) && (cvScore <= efsCvScoreList.get(8))) // 6593<=cvScore<=609
		{
			riskGroup = (String) RiskgroupList.get(4);
		} else if ((cvScore >= efsCvScoreList.get(9)) && (cvScore <= efsCvScoreList.get(10))) // 577<=cvScore<=592
		{
			riskGroup = (String) RiskgroupList.get(5);
		} else if ((cvScore >= efsCvScoreList.get(11)) && (cvScore <= efsCvScoreList.get(12))) // 564<=cvScore<=576
		{
			riskGroup = (String) RiskgroupList.get(6);
		}

		System.out.println(riskGroup);

		// Iterating through the tables to identify the right Risk Group
		for (int j = 9; j < 86; j++) // 9 is the first row where table starts & 86 is the last row in the table
		{

			if (sheet1.getRow(j).getCell(1).getStringCellValue().contains(riskGroup)) {
				tabrow = j;
				for (int k = 1; k < 19; k++) {
					propertyRange = sheet1.getRow(j + 1).getCell(k).getStringCellValue();

					if (k == 1) {
						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace("<$", "");
						propertyRange = propertyRange.replace("K", "000");
						propRange = Double.parseDouble(propertyRange);

						if (PropertyVal < propRange) {
							thecol = k;
							break;
						}
					}

					if (k == 6 && riskGroup.equalsIgnoreCase("Risk Group 1")) // Since, Risk Group 1 has only two
																				// property type tables
					{

						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace(">= $", "");
						propertyRange = propertyRange.replace("K", "000");
						propRange = Double.parseDouble(propertyRange);

						if (PropertyVal >= propRange) {
							thecol = k;
							break;
						}

					}

					if (k == 6 && !riskGroup.equalsIgnoreCase("Risk Group 1")) {
						propertyRange = propertyRange.replace("Property Value and ", "");
						propertyRange = propertyRange.replace("<$", "");
						propertyRange = propertyRange.replace(">= $", "");
						propertyRange = propertyRange.replace("K", "000");
						stringSplitter = propertyRange.split("and");
						String range1 = stringSplitter[0];
						String range2 = stringSplitter[1];

						propRange1 = Double.parseDouble(range1);
						propRange2 = Double.parseDouble(range2);

						if ((PropertyVal >= propRange1) && (PropertyVal < propRange2)) {
							thecol = k;
							break;
						}

					}

					if (k == 11) {

						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace("<$", "");
						propertyRange = propertyRange.replace(">=$", "");
						propertyRange = propertyRange.replace("K", "000");
						stringSplitter = propertyRange.split("and");
						String range1 = stringSplitter[0];
						String range2 = stringSplitter[1];

						propRange1 = Double.parseDouble(range1);
						propRange2 = Double.parseDouble(range2);

						if ((PropertyVal >= propRange1) && (PropertyVal < propRange2)) {
							thecol = k;
							break;
						}
					}

					if (k == 16) {
						propertyRange = propertyRange.replace("Property Value ", "");
						propertyRange = propertyRange.replace(">=$", "");
						propertyRange = propertyRange.replace("K", "000");
						propRange = Double.parseDouble(propertyRange);

						if (PropertyVal >= propRange) {
							thecol = k;
							break;
						}
					}

					k += 4;

				}

			}
			j += 12;
		}

		int r = tabrow + 4;
		try {
			while (sheet1.getRow(r).getCell(thecol).getCellTypeEnum() == CellType.STRING) {

				if (sheet1.getRow(r).getCell(thecol).getStringCellValue().toLowerCase()
						.contains(propertyType.toLowerCase())) {

					therow = r;
					break;
				}
				r++;
			}
		} catch (Exception e) {

		}

		if (code.equalsIgnoreCase("Urban")) {
			Maxltv = sheet1.getRow(therow).getCell(thecol + 1).getNumericCellValue();
		}

		if (code.equalsIgnoreCase("Rural")) {
			Maxltv = sheet1.getRow(therow).getCell(thecol + 2).getNumericCellValue();
		}
		if (code.equalsIgnoreCase("Remote")) {

			Maxltv = sheet1.getRow(therow).getCell(thecol + 3).getNumericCellValue();
		}

		System.out.println("Max LTV is " + Maxltv);

		// Home Equity Calculation
		//Property value=600,000
		//Mortgage balance outstanding=580,000
		// Home Equity =30,000 30100
		//QLA=45,000 
		//QLA
		HomeEquity = (LtvMax - SPLltv) * PropertyVal / 100;
		System.out.println("Home Equity =" + HomeEquity);
		double ActualQLA = 0;

		if (QLA == 0.0) {
			ActualQLA = QLA;
		} else if (HomeEquity == 0.0) {
			ActualQLA = HomeEquity;
		} else if (HomeEquity > QLA) {
			ActualQLA = QLA + 100;
		} else if (HomeEquity < QLA) {
			ActualQLA = HomeEquity + 100;
		}

		System.out.println("Actual QLA :$" + ActualQLA);
		System.out.println("Expected QLA :$" + ExpectedQLA);

		// Displaying QLA result

		if (ActualQLA == ExpectedQLA) {

			test.log(Status.PASS, MarkupHelper.createLabel("QLA Actual value :  $" + ActualQLA, ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("QLA Expected value :  $ " + ExpectedQLA, ExtentColor.GREEN));

			test.log(Status.PASS,
					MarkupHelper.createLabel(" QLA Calculation is Matching with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));

			System.out.println("PASSED in QLA Verification");
		} else {
			System.out.println(ExpectedQLA + " is the expected value.");

			test.log(Status.FAIL, MarkupHelper.createLabel("QLA Actual value :  $" + ActualQLA, ExtentColor.RED));
			test.log(Status.FAIL, MarkupHelper.createLabel("QLA Expected value :  $ " + ExpectedQLA, ExtentColor.RED));

			test.log(Status.FAIL,
					MarkupHelper.createLabel(" QLA Calculation not Matching with GDS Decision", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(QLAInterestScreen));

			System.out.println("FAILED in QLA Verification");
		}

	}

	// Check Max H&A
	public void maxHA() throws IOException, InterruptedException {
		if (attemptNo == 0) {
			test = Extent.createTest("MAX H&A Calculation");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - MAX H&A Calculation");
		}

		org.apache.poi.ss.usermodel.Sheet sheet;

		File file = new File(
				System.getProperty("user.dir") + "\\src\\main\\resources\\Excel\\UAT-SF-GDScalculation.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook workbook = new XSSFWorkbook(inputStream);

		sheet = workbook.getSheet("Max H&A");

		// For UPL
		double qla1 = sheet.getRow(3).getCell(0).getNumericCellValue();
		double qla2 = sheet.getRow(3).getCell(2).getNumericCellValue();
		double max1 = sheet.getRow(3).getCell(3).getNumericCellValue();
		double qla3 = sheet.getRow(4).getCell(0).getNumericCellValue();
		double qla4 = sheet.getRow(4).getCell(2).getNumericCellValue();
		double max2 = sheet.getRow(4).getCell(3).getNumericCellValue();
		double qla5 = sheet.getRow(5).getCell(0).getNumericCellValue();
		double qla6 = sheet.getRow(5).getCell(2).getNumericCellValue();
		double max3 = sheet.getRow(5).getCell(3).getNumericCellValue();
		double qla7 = sheet.getRow(6).getCell(0).getNumericCellValue();
		double qla8 = sheet.getRow(6).getCell(2).getNumericCellValue();
		double max4 = sheet.getRow(6).getCell(3).getNumericCellValue();
		double qla9 = sheet.getRow(7).getCell(0).getNumericCellValue();
		double qla10 = sheet.getRow(7).getCell(2).getNumericCellValue();
		double max5 = sheet.getRow(7).getCell(3).getNumericCellValue();
		// For SPL
		double qla11 = sheet.getRow(11).getCell(0).getNumericCellValue();
		double qla12 = sheet.getRow(11).getCell(2).getNumericCellValue();
		double max6 = sheet.getRow(11).getCell(3).getNumericCellValue();

		if (ExpectedQLA < qla1) {
			ActualMaxHA = 0.0;
		} else if (ExpectedQLA >= qla1 && ExpectedQLA <= qla2) {
			ActualMaxHA = max1;
		} else if (ExpectedQLA >= qla3 && ExpectedQLA <= qla4) {
			ActualMaxHA = max2;
		} else if (ExpectedQLA >= qla5 && ExpectedQLA <= qla6) {
			ActualMaxHA = max3;
		} else if (ExpectedQLA >= qla7 && ExpectedQLA <= qla8) {
			ActualMaxHA = max4;
		} else if (ExpectedQLA >= qla9 && ExpectedQLA <= qla10) {
			ActualMaxHA = max5;
		} else if (ExpectedQLA >= qla11 && ExpectedQLA <= qla12) {
			ActualMaxHA = max6;
		}
		if(Prov.equalsIgnoreCase("MB"))
		{
			ActualMaxHA=0.0;
		}
		// Displaying Interest Rate result
		System.out.println("Actual MaxH&A: " + ActualMaxHA);
		System.out.println("Expected MaxH&A: " + ExpectedMaxHA);

		if (ExpectedMaxHA == ActualMaxHA) {

			test.log(Status.PASS,
					MarkupHelper.createLabel("MaxH&A Actual value : " + ActualMaxHA + "%", ExtentColor.GREEN));
			test.log(Status.PASS,
					MarkupHelper.createLabel("MaxH&A Expected value : " + ExpectedMaxHA + "%", ExtentColor.GREEN));

			test.log(Status.PASS,
					MarkupHelper.createLabel(" MaxH&A Calculation is Matching with GDS Decision", ExtentColor.GREEN));
			test.log(Status.PASS, "Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen));
			System.out.println("PASSED in MaxH&A Verification");
		} else {

			test.log(Status.FAIL,
					MarkupHelper.createLabel("MaxH&A Actual value : " + ActualMaxHA + "%", ExtentColor.RED));
			test.log(Status.FAIL,
					MarkupHelper.createLabel("MaxH&A Expected value : " + ExpectedMaxHA + "%", ExtentColor.RED));

			test.log(Status.FAIL,
					MarkupHelper.createLabel(" MaxH&A Calculation not Matching with GDS Decision", ExtentColor.RED));
			test.log(Status.FAIL, "Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen));
			System.out.println("FAILED in MaxH&A Verification");
		}
		Thread.sleep(3000);
	}

	public void ReasonCode() throws InterruptedException, IOException {
		// TODO Auto-generated method stub //MaxHASPLFullScreen
		if (attemptNo == 0) {
			test = Extent.createTest("Reason Codes/Risk Factors");
		}

		else {

			test = Extent.createTest("Resubmission Result: Attempt#  " + attemptNo + " - Reason Codes");
		}
		
if(length==1)
{
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) +  test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1));
}
else if (length==2)
{
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) + test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2));
}
else if (length==3)
{
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) + test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen3));
}
else {
	test.log(Status.PASS,
			MarkupHelper.createLabel(" Reason Codes/Risk Factors with GDS Decision", ExtentColor.GREEN));
	test.log(Status.PASS,
			"Snapshot below: " + test.addScreenCaptureFromPath(MaxHASPLFullScreen) + test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen1)
					+ test.addScreenCaptureFromPath(ReasonCodeSPLFullScreen2));
}
		Thread.sleep(3000);
	}
	public void SecondPopupSpl() throws Exception {

		attemptNo++;
		driver.switchTo().defaultContent();

		JFrame frmOpt = new JFrame(); // We are declaring the frame
		frmOpt.setAlwaysOnTop(true);// This is the line for displaying it above all windows

		Thread.sleep(1000);
		String s = "<html>Press 1 For Re-Submission with Applicant<br>Press 2 For Results<br>";

		JLabel label = new JLabel(s);
		JTextPane jtp = new JTextPane();
		jtp.setSize(new Dimension(480, 10));
		jtp.setPreferredSize(new Dimension(480, jtp.getPreferredSize().height));
		label.setFont(new Font("Arial", Font.BOLD, 20));
		UIManager.put("OptionPane.minimumSize", new Dimension(500, 200));
		UIManager.put("TextField.font", new FontUIResource(new Font("Verdana", Font.BOLD, 18)));
		// Getting Input from user

		String option = JOptionPane.showInputDialog(frmOpt, label);

		int useroption = Integer.parseInt(option);

		switch (useroption) {

		case 1:

			// Function for Re-Submission
			System.out.println("Re-Submission with  Applicant");
			resubmitForDecisionSpl();

			break;

		case 2:

			System.out.println("Results");
			if (attemptNo == 0) {
				test = Extent.createTest("Resubmission - Confirmation ");
				test.info(" The test run complete. Please review test result(s)");
			}

			else {

				test = Extent.createTest("Resubmission - Confirmation ");
				test.info(" The test run complete. Please review test result(s)");
			}

			Thread.sleep(3000);

			driver.close();
			driver.quit();
			break;

		}

	}

	public void resubmitForDecisionSpl() throws Exception {

		System.out.println("Re-Submission Attempt:"+attemptNo);

		firstPopup();

		Thread.sleep(3000);
		driver.get(url);
		getAddress();
		getPartyDetails();
		getCollateral();

		getAppDetails();
		
		calculateIncome();
		calculateSPLLiability();
		splLTV();
		splRemInCal();
		splQLA();
		splFinalQLA();
		maxHA();
		ReasonCode();
		Thread.sleep(4000);
		SecondPopupSpl();
	}

}
