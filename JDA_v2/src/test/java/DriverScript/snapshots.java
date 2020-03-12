package DriverScript;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.Properties;

import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import Functions.DataBase;
import Functions.DififoReportSetup;
import Functions.ExcelFile;

public class snapshots extends DififoReportSetup{
	
	private static final int i = 0;
	String ItemNumber;
	String LocationNum;
	String testScenarioFilePath;
	String testCaseFileName;
	String testdatasheet;
	String Externaldata;
	String InputfromWIP;
	String inputDatafromDBFileName;
	String DFUSheet;
	String DDESheet;
	String DFUMOVINGEVENTMAP;
	String DFUDDEMAP;
	String DFUEFFPRICE;
	String FCSTnew;
	String lewandowskiparam;
	String meanvalueadj;
	String targetdfusheet;
	
	String WIPdata;
	String DDEProfiledata;
	String DFUData;
	String DFUMOVINGData;
	String DFUDDEData;
	String DFUEFFData;
	String FCSTData;
	String LEWANDData;
	String MeanvalueData;
	String Snapshotwb;
	String snapshotpath;
	ExcelFile exfile = new ExcelFile();
	
	public String fetchItemNumber(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException
	{
    	ItemNumber = exfile.readExcel(filePa, fileNa, SheetNa, row, 0);
    	return ItemNumber;
	}
	
	public String fetchLocationNumber(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	LocationNum = exfile.readExcel(filePa, fileNa, SheetNa, row, 2);
    	return LocationNum;
            }
	
	/*@BeforeTest
	public void excelFileClear() throws IOException, InterruptedException
	{
		try {
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
				
		testScenarioFilePath = envProp.getProperty("testScenarioFilePath");
		inputDatafromDBFileName = envProp.getProperty("snapshotsInputFile");
		DDEProfiledata = envProp.getProperty("DDE_Sheet");
		
		DataBase db= new DataBase();
		db.cleanSheet(DDEProfiledata);
		
		
		}
		catch(Exception e) {
			System.out.println(e.getMessage());
			return;
		}
	}*/
		
	
	
	@Test(priority=0)
	public void DDEPROFILE_Validation() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		DDESheet = envProp.getProperty("DDE_Sheet");
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,InputfromWIP,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,InputfromWIP,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String DDEquery = queryProp.getProperty("DDE");
		String DDESUPSQuery = queryProp.getProperty("DDE_SUPS");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String DDEloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.DFU_DDE_PROFILE(DDEquery + DMDunit +"' and LOC ='" + DDEloc +"'"+" order by DDEPROFILEID");
		db.DFU_DDE_SUPSValid(DDESUPSQuery + DMDunit +"' and LOC ='" + DDEloc +"'"+" order by DDEPROFILEID");
		db.minusquery1(DDEquery + DMDunit +"' and LOC ='" + DDEloc +"'" +" minus "+ DDESUPSQuery + DMDunit +"' and LOC ='" + DDEloc +"'",DDESheet);
		
		}
		
	}
	
	@Test(priority=1)
	public void DFU_Validation() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		DFUSheet = envProp.getProperty("DFU_Sheet");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String DFUQuery = queryProp.getProperty("DFU");
		String DFUSUPSQuery = queryProp.getProperty("SUPS_DFU");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String DDEloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.DFU_Validation(DFUQuery + DMDunit +"' and LOC ='" + DDEloc +"'");
		db.DFU_SUPSValidation(DFUSUPSQuery + DMDunit +"' and LOC ='" + DDEloc +"'");
		db.minusquery1(DFUQuery + DMDunit +"' and LOC ='" + DDEloc +"'" +" minus "+ DFUSUPSQuery + DMDunit +"' and LOC ='" + DDEloc +"'",DFUSheet);
		}
	}
	
	@Test(priority=2)
	public void DFUMOVINGEVENTMAP() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		DFUMOVINGEVENTMAP = envProp.getProperty("DFUMOVINGEVENTMAP");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String EVENTMAP = queryProp.getProperty("EVENTMAP");
		String SUPSEVENTMAP = queryProp.getProperty("SUPSEVENTMAP");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String eventloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.Eventmap_Validation(EVENTMAP + DMDunit +"' and LOC ='" + eventloc +"'");
		db.Eventmap_SUPSValidation(SUPSEVENTMAP + DMDunit +"' and LOC ='" + eventloc +"'");
		db.minusquery1(EVENTMAP + DMDunit +"' and LOC ='" + eventloc +"'" +" minus "+ SUPSEVENTMAP + DMDunit +"' and LOC ='" + eventloc +"'",DFUMOVINGEVENTMAP);
		}
	}
	
	@Test(priority=3)
	public void DFUDDEMAP() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		DFUDDEMAP = envProp.getProperty("DFUDDEMAP_Sheet");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String DFUDDE = queryProp.getProperty("DFUDDEMAP");
		String SUPSDFUDDE = queryProp.getProperty("SUPSDFUDDEMAP");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String eventloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.DFUDDE_Validation(DFUDDE + DMDunit +"' and LOC ='" + eventloc +"'");
		db.DFUDDE_SUPSValidation(SUPSDFUDDE + DMDunit +"' and LOC ='" + eventloc +"'");
		db.minusquery1(DFUDDE + DMDunit +"' and LOC ='" + eventloc +"'" +" minus "+ SUPSDFUDDE + DMDunit +"' and LOC ='" + eventloc +"'",DFUDDEMAP);
		}
	}
	
	@Test(priority=4)
	public void DFUEFFPRICE() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		DFUEFFPRICE = envProp.getProperty("DFUEFF_Sheet");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String DFUEFF = queryProp.getProperty("DFUEFFPRICE");
		String SUPSDFUEFF = queryProp.getProperty("SUPSDFUEFFPRICE");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String eventloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.DFUEFF_Validation(DFUEFF + DMDunit +"' and LOC ='" + eventloc +"'");
		db.DFUEFF_SUPSValidation(SUPSDFUEFF + DMDunit +"' and LOC ='" + eventloc +"'");
		db.minusquery1(DFUEFF + DMDunit +"' and LOC ='" + eventloc +"'" +" minus "+ SUPSDFUEFF + DMDunit +"' and LOC ='" + eventloc +"'",DFUEFFPRICE);
		}
	}
	
	@Test(priority=5)
	public void FCST() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		FCSTnew = envProp.getProperty("FCST_Sheet");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String FCST = queryProp.getProperty("FCST");
		String SUPSFCST = queryProp.getProperty("SUPSFCST");
		String DMDunitparent = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String eventloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.FCST_Validation(FCST + DMDunitparent +"' and LOC ='" + eventloc +"'");
		db.FCST_SUPSValidation(SUPSFCST + DMDunitparent +"' and LOC ='" + eventloc +"'");
		db.minusquery1(FCST + DMDunitparent +"' and LOC ='" + eventloc +"'" +" minus "+ SUPSFCST + DMDunitparent +"' and LOC ='" + eventloc +"'",FCSTnew);
		
		String DMDunitchild = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,1);
		db.FCST_Validation(FCST + DMDunitchild +"' and LOC ='" + eventloc +"'");
		db.FCST_SUPSValidation(SUPSFCST + DMDunitchild +"' and LOC ='" + eventloc +"'");
		db.minusquery1(FCST + DMDunitchild +"' and LOC ='" + eventloc +"'" +" minus "+ SUPSFCST + DMDunitchild +"' and LOC ='" + eventloc +"'",FCSTnew);
		}
	}
	
	@Test(priority=6)
	public void lewandowskiparam() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		lewandowskiparam = envProp.getProperty("Lewand_Sheet");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String LEWAND = queryProp.getProperty("LEWAND");
		String SUPSLEWAND = queryProp.getProperty("SUPSLEWAND");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String eventloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.LEWAND_Validation(LEWAND + DMDunit +"' and LOC ='" + eventloc +"'");
		db.LEWAND_SUPSValidation(SUPSLEWAND + DMDunit +"' and LOC ='" + eventloc +"'");
		db.minusquery1(LEWAND + DMDunit +"' and LOC ='" + eventloc +"'" +" minus "+ SUPSLEWAND + DMDunit +"' and LOC ='" + eventloc +"'",lewandowskiparam);
		}
	}
	
	@Test(priority=7)
	public void meanvalueadj() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		meanvalueadj = envProp.getProperty("Meanvalue_Sheet");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String Meanvalue = queryProp.getProperty("Meanvalue");
		String SUPSMeanvalue = queryProp.getProperty("SUPSMeanvalue");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		String eventloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.Meanvalue_Validation(Meanvalue + DMDunit +"' and LOC ='" + eventloc +"'");
		db.Meanvalue_SUPSValidation(SUPSMeanvalue + DMDunit +"' and LOC ='" + eventloc +"'");
		db.minusquery1(Meanvalue + DMDunit +"' and LOC ='" + eventloc +"'" +" minus "+ SUPSMeanvalue + DMDunit +"' and LOC ='" + eventloc +"'",meanvalueadj);
		}
	}
	
	@Test(priority=8)
	public void targetdfumap() throws IOException, SQLException, InterruptedException 
	{
		
		InputStream envPropInput = new FileInputStream("./Environment\\Environment.properties");
		Properties envProp = new Properties();
		envProp.load(envPropInput);
		
		testScenarioFilePath = envProp.getProperty("snapshotstestDataFilePath");
		testCaseFileName = envProp.getProperty("snapshotsInputFile");
		testdatasheet = envProp.getProperty("Input_Sheet");
		InputfromWIP = envProp.getProperty("WIP_Sheet");
		targetdfusheet = envProp.getProperty("target_sheet");		
		int rowMax = exfile.getTotalRowColumn(testScenarioFilePath,testCaseFileName,testdatasheet);
		
		report.log("No of Test data provided by the user for validating Snapshots: "+ rowMax );
		for (int i=1;i<=rowMax;i++)
		{
		//String FromItem = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,0);
		//String LOC = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		
		DataBase db= new DataBase();
		InputStream queryPropInput = new FileInputStream("./DB Query\\Query1.properties");
		Properties queryProp = new Properties();
		queryProp.load(queryPropInput);
		//String Query = queryProp.getProperty("WIP_SAP_SUPERSESSION_NEW");
		//db.WIP_SAP_SUPERSESSION(Query + FromItem +"' and LOC ='" + LOC +"'");
		
		String targetdfu = queryProp.getProperty("targetdfu");
		String SUPStargetdfu = queryProp.getProperty("SUPStargetdfu");
		String DMDunit = fetchItemNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,1);
		String eventloc = fetchLocationNumber(testScenarioFilePath,testCaseFileName,testdatasheet,i,2);
		db.targetdfu_Validation(targetdfu + DMDunit +"' and LOC ='" + eventloc +"'");
		db.targetdfu_SUPSValidation(SUPStargetdfu + DMDunit +"' and LOC ='" + eventloc +"'");
		db.minusquery1(targetdfu + DMDunit +"' and LOC ='" + eventloc +"'" +" minus "+ SUPStargetdfu + DMDunit +"' and LOC ='" + eventloc +"'",targetdfusheet);
		}
	}
}
