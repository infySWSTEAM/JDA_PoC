///************************************************************************************************************************
//		Author           : SGWS JDA Team 
//		Last Modified by : Prema Kothandaraman
//		Last Modified on : 13-Feb-2020
//		Summary 		 : JDA UI Automation scripts for pages Demand Workbench, PIM_Supersession and Supersession
//
//************************************************************************************************************************/

package DriverScript;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.Properties;

import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import Functions.DififoReportSetup;
import Functions.ExcelFile;
import ObjectRepository.DemandWorkBench;
import ObjectRepository.HomePage;
import ObjectRepository.LoginPage;
import ObjectRepository.PIMSupersession;
import ObjectRepository.Supersession;

public class JDA extends DififoReportSetup{
	
	WebDriver driver = null;
	ExcelFile exf = null;
	
	String url;
	String testScenarioFilePath;
	String testCaseFileName;
	String testCaseSheetName;
	String testcasename;
	String testdatasheet;
	String query;
	
		
	public void Screenshot(String Filename) throws IOException
	{
		File file=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
		report.addImage(file, Filename);
	}	
	
	@BeforeMethod
	public void beforeTest() throws SQLException, Exception {
		System.setProperty("webdriver.chrome.driver", "./Drivers\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().deleteAllCookies();
		Thread.sleep(3000);
		exf = new ExcelFile();
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		url= envProp.getProperty("URL");
		int row=1;
		testScenarioFilePath = envProp.getProperty("testScenarioFilePath");
		testCaseFileName = envProp.getProperty("testCaseFileName");
		testCaseSheetName = envProp.getProperty("testCaseSheetName");
		testcasename = envProp.getProperty("testcasename");
		testdatasheet = envProp.getProperty("testdatasheet");
		driver.get(url);
		Thread.sleep(3000);
		driver.manage().window().maximize();
		driver.switchTo().activeElement();
		LoginPage lp= new LoginPage(driver);
		Thread.sleep(1000);
		report.log("Inputting the Username and Password");
		lp.setUsername(testScenarioFilePath,testCaseFileName,testCaseSheetName,row,0);
		Thread.sleep(1000);
		lp.setPassword(testScenarioFilePath,testCaseFileName,testCaseSheetName,row,1);
		Thread.sleep(2000);
		Screenshot("Login Entered");
		lp.clickSubmit();
		report.log("User logged in successfully");
		Thread.sleep(2000);
		report.log("User is viewing JDA Homepage now");
		Screenshot("HomePage");
		Thread.sleep(5000);
	}
	
	@Test(priority=2)
	public void JDASupersession() throws SQLException, Exception{
		exf = new ExcelFile();
		int rowMax = exf.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		for (int i=1;i<=rowMax;i++)
		{
		HomePage hp= new HomePage(driver);
		hp.setsearchname(testScenarioFilePath,testCaseFileName,testcasename,1,0);
		Thread.sleep(5000);
		report.log("Searching for Supersession UI by inputting the page name");
		report.log("Identifying from the Dropdown list");
		Screenshot("Inputting the search name");
		Supersession ss = new Supersession(driver);
		ss.dropdownclick();
		Thread.sleep(10000);
		driver.switchTo().frame("appFrame");
		ss.Clear();
	    driver.switchTo().frame("PromptScreenPopupFrame");
	    ss.setInputItem(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		ss.setInputLoc(testScenarioFilePath,testCaseFileName,testdatasheet,i,1);
		report.log("Inputting the Item and Location details");
		Screenshot("search criteria");
		driver.switchTo().defaultContent();
		driver.switchTo().frame("appFrame");
		ss.clickdone();
		report.log("Search Results from UI");
		Thread.sleep(10000);
		driver.switchTo().defaultContent();
		Thread.sleep(5000);
		Screenshot("Supersession records in UI");
		Thread.sleep(5000);
		hp.clr();
	}
}

	
	@Test(priority=1)
	public void JDAPIMSupersession() throws SQLException, Exception{
		
		exf = new ExcelFile();
		int rowMax = exf.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		for (int i=1;i<=rowMax;i++)
		{
		HomePage hp= new HomePage(driver);
		hp.setsearchname(testScenarioFilePath,testCaseFileName,testcasename,2,0);
		report.log("Searching for PIM Supersession UI by inputting the page name");
		report.log("Identifying from the Dropdown list");
		Screenshot("Inputting the search name");
		PIMSupersession ps = new PIMSupersession(driver);
		ps.dropdownclick();
		Thread.sleep(10000);
		driver.switchTo().frame("appFrame");
		ps.Clear();
	    driver.switchTo().frame("PromptScreenPopupFrame");
	    ps.setUltimateParent(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
	    report.log("Inputting the Item and Location details");
		Screenshot("search criteria");
	    driver.switchTo().defaultContent();
		driver.switchTo().frame("appFrame");
		ps.clickdone();
		report.log("Search Results from UI");
		Thread.sleep(10000);  
		driver.switchTo().defaultContent();
		Screenshot("PIM_Supersession records in UI");
		Thread.sleep(2000);  
		hp.clr();
		}
	}
	
	@Test(priority=1)
	public void JDADemandWorkBench() throws SQLException, Exception{
		
		exf = new ExcelFile();
		int rowMax = exf.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		for (int i=1;i<=rowMax;i++)
		{
		HomePage hp= new HomePage(driver);
		hp.setsearchname(testScenarioFilePath,testCaseFileName,testcasename,3,0);
		report.log("Searching for Demand workbench UI by inputting the page name");
		report.log("Identifying from the Dropdown list");
		Screenshot("Inputting the search name");
		DemandWorkBench wb = new DemandWorkBench(driver);
		wb.dropdownclick();
		Thread.sleep(7000);
		driver.switchTo().frame("appFrame");
		Thread.sleep(2000);
		wb.Clear();
	    driver.switchTo().frame("PromptScreenPopupFrame");
		Thread.sleep(6000);
	    wb.setDMDUNIT(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
	    wb.setLOC(testScenarioFilePath,testCaseFileName,testdatasheet,i,1);
	    wb.setDMDGROUP(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
	    report.log("Inputting the Dmdunit, Dmdgroup and Location details");
		Screenshot("search criteria");
	    driver.switchTo().defaultContent();
		driver.switchTo().frame("appFrame");
		wb.clickdone();
		report.log("Search Results from UI");
		Thread.sleep(18000);
		driver.switchTo().defaultContent();
		Screenshot("Demand Workbench is now visible");
		Thread.sleep(2000);
		hp.clr();
	}
			}

		
	@AfterMethod
	public void JDALogout() throws InterruptedException, IOException, SQLException
	{
		LoginPage lp= new LoginPage(driver);
		Thread.sleep(2000);
		lp.useroption();
		Thread.sleep(2000);
		report.log("User trying to check the Logout options");
		Screenshot("User logout options");
		lp.JDAlogout();
		Thread.sleep(2000);
		report.log("User logging out of JDA UI");
		Screenshot("User Logged out");
		driver.quit();
		Thread.sleep(2000);
}
	}