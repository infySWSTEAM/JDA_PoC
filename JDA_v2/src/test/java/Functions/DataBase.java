///************************************************************************************************************************
//		Author           : SGWS JDA Team 
//		Last Modified by : Anushya Karunakaran
//		Last Modified on : 13-Feb-2020
//		Summary 		 : SQL Validations for SS Classification Rejection scenarios
//
//************************************************************************************************************************/

package Functions;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataBase extends DififoReportSetup {
	
	XSSFWorkbook workbook  = null;
	String testDataFilePath;
	String snapShotFilePath;
	String connectionURL;
	String userName;
	String pass;
	
	public void dbopen() throws IOException, SQLException {
		
	InputStream input = new FileInputStream("./Environment\\Environment.properties");
	Properties prop = new Properties();
	prop.load(input);
	testDataFilePath = prop.getProperty("testDataFilePath")+"\\"+prop.getProperty("inputDatafromDBFileName");
	connectionURL=prop.getProperty("ConnectionURL");
	userName=prop.getProperty("DBUserName");
	pass= prop.getProperty("DBPass");
	
	File  file = new File(testDataFilePath);
	FileInputStream inputfile = new FileInputStream(file);
	 workbook = new XSSFWorkbook(inputfile);
	 		}	
	
	public void dbopenSnap() throws IOException, SQLException {
		
		InputStream input = new FileInputStream("./Environment\\Environment.properties");
		Properties prop = new Properties();
		prop.load(input);
		
		snapShotFilePath = prop.getProperty("testDataFilePath")+"\\"+prop.getProperty("snapshotsInputFile");
		connectionURL=prop.getProperty("ConnectionURL");
		userName=prop.getProperty("DBUserName");
		pass= prop.getProperty("DBPass");
		
		File  file = new File(snapShotFilePath);
		FileInputStream inputfile = new FileInputStream(file);
		 workbook = new XSSFWorkbook(inputfile);
		 		}	
	public void cleanSheet(String ipsheet) {
		
		Sheet sheet1 = workbook.getSheet(ipsheet);
		
		int noOfRows = sheet1.getPhysicalNumberOfRows();
		
		if (noOfRows > 0) {
			for(int i = sheet1.getFirstRowNum(); i<= sheet1.getLastRowNum(); i++) {
				if(sheet1.getRow(i)!= null) {
					sheet1.removeRow(sheet1.getRow(i));
				} else {
					System.out.println("Info: clean sheet='");
				}
			}
		} 
	}
	
	public void dbJDATestInput(String Query) throws SQLException, IOException {
		
		
		InputStream input = new FileInputStream("./Environment\\Environment.properties");
		Properties prop = new Properties();
		prop.load(input);
		testDataFilePath = prop.getProperty("testDataFilePath")+"\\"+prop.getProperty("testCaseFileName");
		connectionURL=prop.getProperty("ConnectionURL");
		userName=prop.getProperty("DBUserName");
		pass= prop.getProperty("DBPass");
		
		File  file = new File(testDataFilePath);
		FileInputStream inputfile = new FileInputStream(file);
		workbook = new XSSFWorkbook(inputfile);
		 
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
		Statement stmt=con.createStatement();
		
		report.log(Query);
		
		ResultSet resSet = stmt.executeQuery(Query);
				
		Sheet sheet1 = workbook.getSheet("TestData");
		System.out.println(sheet1);
		Row rowhead = sheet1.createRow((short) 0);
		rowhead.createCell((short) 0).setCellValue("Item");
		rowhead.createCell((short) 1).setCellValue("Location");
		rowhead.createCell((short) 2).setCellValue("Dmdgrp");
		
		
		int rowmax=sheet1.getLastRowNum();
		int i=rowmax+1;
		while (resSet.next())
		{		
			Row row = sheet1.createRow((short) i++);
			row.createCell((short) 0).setCellValue(resSet.getString("Item"));
			row.createCell((short) 1).setCellValue(resSet.getString("Location"));
			row.createCell((short) 2).setCellValue(resSet.getString("Dmdgrp"));
							
			}
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);
		con.close();
	}
	
	public void dbPIMSupersessionConn(String Query) throws SQLException, IOException
	{
		dbopen();
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
		Statement stmt=con.createStatement();
		
		report.log(Query);
		
		ResultSet resSet = stmt.executeQuery(Query);
				
		Sheet sheet1 = workbook.getSheet("PIMSupersession");
		
		Row rowhead = sheet1.createRow((short) 0);
		rowhead.createCell((short) 0).setCellValue("SGWS_ITEM_NUMBER");
		rowhead.createCell((short) 1).setCellValue("REPLACED_BY_ID");
		rowhead.createCell((short) 2).setCellValue("REPLACEMENT_ITEM_EFF_DT");
		rowhead.createCell((short) 3).setCellValue("ULTIMATE_PARENT");
		rowhead.createCell((short) 4).setCellValue("LEVEL_ID");
		rowhead.createCell((short) 5).setCellValue("RUN_DATE");
		
		int rowmax=sheet1.getLastRowNum();
		int i=rowmax+1;
		while (resSet.next())
		{		
			Row row = sheet1.createRow((short) i++);
			row.createCell((short) 0).setCellValue(resSet.getString("SGWS_ITEM_NUMBER"));
			row.createCell((short) 1).setCellValue(resSet.getString("REPLACED_BY_ID"));
			row.createCell((short) 2).setCellValue(resSet.getString("REPLACEMENT_ITEM_EFF_DT"));
			row.createCell((short) 3).setCellValue(resSet.getString("ULTIMATE_PARENT"));
			row.createCell((short) 4).setCellValue(resSet.getString("LEVEL_ID"));
			row.createCell((short) 5).setCellValue(resSet.getString("RUN_DATE"));
					
			}
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);
		con.close();
	}
	
	public void dbExtSupersessionConn(String Query) throws SQLException, IOException
	{
		dbopen();
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
		Statement stmt=con.createStatement();
		report.log(Query);
		ResultSet resSet = stmt.executeQuery(Query);
		Sheet sheet2 = workbook.getSheet("ExtSupersession");
		
		Row rowhead1 = sheet2.createRow((short) 0);
		rowhead1.createCell((short) 0).setCellValue("OLD_ITEM");
		rowhead1.createCell((short) 1).setCellValue("LOC");
		rowhead1.createCell((short) 2).setCellValue("NEW_ITEM");
		rowhead1.createCell((short) 3).setCellValue("EFF_DATE");
		rowhead1.createCell((short) 4).setCellValue("ULT_PARENT");
		rowhead1.createCell((short) 5).setCellValue("SITE");
			
		int rowmax=sheet2.getLastRowNum();
		int i=rowmax+1;
		while (resSet.next())
		{		
			Row row = sheet2.createRow((short) i++);
			row.createCell((short) 0).setCellValue(resSet.getString("OLD_ITEM"));
			row.createCell((short) 1).setCellValue(resSet.getString("LOC"));
			row.createCell((short) 2).setCellValue(resSet.getString("NEW_ITEM"));
			row.createCell((short) 3).setCellValue(resSet.getString("EFF_DATE"));
			row.createCell((short) 4).setCellValue(resSet.getString("ULT_PARENT"));
			row.createCell((short) 5).setCellValue(resSet.getString("SITE"));
					
			}
				
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);
		con.close();
	}
	public void oldSkuCheck(String Query,String OldItem, int i) throws IOException, SQLException 
	{
		dbopen();
		
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
		Statement stmt=con.createStatement();
		report.log(Query);
		ResultSet resSet = stmt.executeQuery(Query);
		Sheet sheet2 = workbook.getSheet("ExtSupersession");
		Row rowhead1 = sheet2.getRow(0);
		Cell cell = rowhead1.createCell((short) 6);
		cell.setCellValue("Old Item check");
		if (resSet.next()==false)
		{			
			Row row = sheet2.getRow((short) i);
			row.createCell((short) 6).setCellValue("Old Item/Loc not present in table SKU,");
		}
		else
		{
			Row row = sheet2.getRow((short) i);
			row.createCell((short) 6).setCellValue("");
						
		}
		
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);
		con.close();
	}
	
	public void newSkuCheck(String Query,String NewItem, int i) throws IOException, SQLException, InterruptedException 
	{
		dbopen();
		Thread.sleep(5000);
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
		Thread.sleep(5000);
		Statement stmt=con.createStatement();
		report.log(Query);
		ResultSet resSet = stmt.executeQuery(Query);
		Sheet sheet2 = workbook.getSheet("ExtSupersession");
		Row rowhead1 = sheet2.getRow(0);
		Cell cell1 = rowhead1.createCell((short) 7);
		cell1.setCellValue("New Item check");
		
		if (resSet.next()==false)
		{			
			Row row = sheet2.getRow((short) i);
			row.createCell((short) 7).setCellValue("New Item/Loc not present in table SKU,");
		}
		else 
		{
			Row row = sheet2.getRow((short) i);
			row.createCell((short) 7).setCellValue("");
		}
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);	
		con.close();
	}
	
	public void cellMerge(int i) throws IOException
	{   		
		try {
		Sheet sheet2 = workbook.getSheet("ExtSupersession");
		Row rowhead1 = sheet2.getRow(0);
		Cell cell1 = rowhead1.createCell((short) 8);
		cell1.setCellValue("Merged Cell");
		Row row = sheet2.getRow((short) i);
		
		Cell olditem = row.getCell(6);
		String a=olditem.getStringCellValue();
		
		Cell newitem = row.getCell(7);
		String b=newitem.getStringCellValue();

		Cell merge = row.createCell(8);

		merge.setCellValue(a+b);
		}
		catch (Exception e) {
            return;
        }
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);

		
	}
	public void dbSkuRejection(String Query) throws SQLException, IOException, InterruptedException
	{
		dbopen();
		Thread.sleep(5000);
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
		Thread.sleep(5000);
		Statement stmt=con.createStatement();
		report.log(Query);
		ResultSet resSet = stmt.executeQuery(Query);
		Sheet sheet3 = workbook.getSheet("SKU_Rejection");
		
		Row rowhead1 = sheet3.createRow((short) 0);
		rowhead1.createCell((short) 0).setCellValue("FROMITEM");
		rowhead1.createCell((short) 1).setCellValue("TOITEM");
		rowhead1.createCell((short) 2).setCellValue("LOC");
		rowhead1.createCell((short) 3).setCellValue("REJECT_REASON");
		rowhead1.createCell((short) 4).setCellValue("RUN_DATE");
					
		int rowmax=sheet3.getLastRowNum();
		int i=rowmax+1;
		while (resSet.next())
		{		
			Row row = sheet3.createRow((short) i++);
			row.createCell((short) 0).setCellValue(resSet.getString("FROMITEM"));
			row.createCell((short) 1).setCellValue(resSet.getString("TOITEM"));
			row.createCell((short) 2).setCellValue(resSet.getString("LOC"));
			row.createCell((short) 3).setCellValue(resSet.getString("REJECT_REASON"));
			row.createCell((short) 4).setCellValue(resSet.getString("RUN_DATE"));
			}
				
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);
		con.close();
	}
	
	public void rejectioncompare(int i,int j) throws IOException, SQLException
	{	
		
		dbopen();
		Sheet sheet3=workbook.getSheet("SKU_Rejection");
		Row row0=sheet3.getRow(0);
		Cell cell5=row0.createCell(5);
		cell5.setCellValue("Compared result");
	
		Sheet sheet2= workbook.getSheet("ExtSupersession");
		
		Cell extold1 = sheet2.getRow(i).getCell(0);
		String extold = extold1.getStringCellValue();
				
		Cell extnew1 = sheet2.getRow(i).getCell(2);
		String extnew = extnew1.getStringCellValue();
				
		Cell extloc1=sheet2.getRow(i).getCell(1);
		String extloc = extloc1.getStringCellValue();
			
		Cell der_rsn1=sheet2.getRow(i).getCell(8);
		String der_rsn = der_rsn1.getStringCellValue();
		
		
		Cell skuold1 = sheet3.getRow(j).getCell(0);
		String skuold = skuold1.getStringCellValue();
				
		Cell skunew1 = sheet3.getRow(j).getCell(1);
		String skunew = skunew1.getStringCellValue();
			
		Cell skuloc1=sheet3.getRow(j).getCell(2);
		String skuloc = skuloc1.getStringCellValue();
				
		Cell rej_rsn1=sheet3.getRow(j).getCell(3);
		String rej_rsn = rej_rsn1.getStringCellValue();
		
		
		if (der_rsn.isEmpty()) 
		{
			Row row = sheet2.getRow((short) i);
			row.createCell((short) 9).setCellValue("VALID SKUs");
		}
		else if ((extold.equals(skuold)) && (extnew.equals(skunew)) && (extloc.equals(skuloc)) && (der_rsn.equals(rej_rsn)))
	    {
			Row row = sheet3.getRow((short) j);
			row.createCell((short) 5).setCellValue("PASS");
			Row row1 = sheet2.getRow((short) i);
			row1.createCell((short) 9).setCellValue("Rejections validated PASSED");
		}
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);
		
	}
				
		public void invalidData(int i,int j) throws IOException, SQLException
		{	
					
			dbopen();
			DataFormatter formatter = new DataFormatter();
			Sheet sheet2= workbook.getSheet("ExtSupersession");
			Row row0=sheet2.getRow(0);
			Cell cell5=row0.createCell(9);
			cell5.setCellValue("Validation result");
		
			Cell err1=sheet2.getRow(i).getCell(9);
			String cellValue = formatter.formatCellValue(err1);
						
			if (cellValue.isEmpty())
		{
			Row row1 = sheet2.getRow((short) i);
			row1.createCell((short) 9).setCellValue("Sourcing missing,Invalid Supersessions");
		}	
						
		FileOutputStream fileOut = new FileOutputStream(testDataFilePath);
		workbook.write(fileOut);
		}
		
		
		public void WIP_SAP_SUPERSESSION(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("WIP_SAP_SUPERSESSION_NEW");
			
			Row rowhead1 = sheet1.createRow((short) 0);
			rowhead1.createCell((short) 0).setCellValue("FROMITEM");
			rowhead1.createCell((short) 1).setCellValue("TOITEM");
			rowhead1.createCell((short) 2).setCellValue("LOC");
			rowhead1.createCell((short) 3).setCellValue("REC_TYPE");
							
			int rowmax=sheet1.getLastRowNum();
			int i=rowmax+1;
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("FROMITEM"));
				row.createCell((short) 1).setCellValue(resSet.getString("TOITEM"));
				row.createCell((short) 2).setCellValue(resSet.getString("LOC"));
				row.createCell((short) 3).setCellValue(resSet.getString("REC_TYPE"));
									
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		
		public void DFU_DDE_PROFILE(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DDE_Profile");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowhead = rowmax+5;
			
			if (rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("DDEPROFILEID");
			rowhead1.createCell((short) 2).setCellValue("STARTPCT1");
			rowhead1.createCell((short) 3).setCellValue("STARTPCT2");
			rowhead1.createCell((short) 4).setCellValue("STARTPCT3");
			rowhead1.createCell((short) 5).setCellValue("STARTPCT4");
			rowhead1.createCell((short) 6).setCellValue("STARTPCT5");
			rowhead1.createCell((short) 7).setCellValue("STARTPCT6");
			rowhead1.createCell((short) 8).setCellValue("DMDCAL");
			rowhead1.createCell((short) 9).setCellValue("STARTDATE");
			rowhead1.createCell((short) 10).setCellValue("DESCR");
			rowhead1.createCell((short) 11).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 12).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 13).setCellValue("LOC");
			}
					
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowhead);
				rowhead1.createCell((short) 0).setCellValue("MODEL");
				rowhead1.createCell((short) 1).setCellValue("DDEPROFILEID");
				rowhead1.createCell((short) 2).setCellValue("STARTPCT1");
				rowhead1.createCell((short) 3).setCellValue("STARTPCT2");
				rowhead1.createCell((short) 4).setCellValue("STARTPCT3");
				rowhead1.createCell((short) 5).setCellValue("STARTPCT4");
				rowhead1.createCell((short) 6).setCellValue("STARTPCT5");
				rowhead1.createCell((short) 7).setCellValue("STARTPCT6");
				rowhead1.createCell((short) 8).setCellValue("DMDCAL");
				rowhead1.createCell((short) 9).setCellValue("STARTDATE");
				rowhead1.createCell((short) 10).setCellValue("DESCR");
				rowhead1.createCell((short) 11).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 12).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 13).setCellValue("LOC");
			}
			
			if(rowmax==0) {
			int i=rowmax+1;
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("DDEPROFILEID"));
				row.createCell((short) 2).setCellValue(resSet.getString("STARTPCT1"));
				row.createCell((short) 3).setCellValue(resSet.getString("STARTPCT2"));
				row.createCell((short) 4).setCellValue(resSet.getString("STARTPCT3"));
				row.createCell((short) 5).setCellValue(resSet.getString("STARTPCT4"));
				row.createCell((short) 6).setCellValue(resSet.getString("STARTPCT5"));
				row.createCell((short) 7).setCellValue(resSet.getString("STARTPCT6"));
				row.createCell((short) 8).setCellValue(resSet.getString("DMDCAL"));
				row.createCell((short) 9).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 10).setCellValue(resSet.getString("DESCR"));
				row.createCell((short) 11).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 12).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 13).setCellValue(resSet.getString("LOC"));
			
				}
			}
			else {
				int i=newrowhead+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
					row.createCell((short) 1).setCellValue(resSet.getString("DDEPROFILEID"));
					row.createCell((short) 2).setCellValue(resSet.getString("STARTPCT1"));
					row.createCell((short) 3).setCellValue(resSet.getString("STARTPCT2"));
					row.createCell((short) 4).setCellValue(resSet.getString("STARTPCT3"));
					row.createCell((short) 5).setCellValue(resSet.getString("STARTPCT4"));
					row.createCell((short) 6).setCellValue(resSet.getString("STARTPCT5"));
					row.createCell((short) 7).setCellValue(resSet.getString("STARTPCT6"));
					row.createCell((short) 8).setCellValue(resSet.getString("DMDCAL"));
					row.createCell((short) 9).setCellValue(resSet.getString("STARTDATE"));
					row.createCell((short) 10).setCellValue(resSet.getString("DESCR"));
					row.createCell((short) 11).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 12).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 13).setCellValue(resSet.getString("LOC"));
				
					}
			}
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void DFU_DDE_SUPSValid(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DDE_Profile");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+3;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("DDEPROFILEID");
			rowhead1.createCell((short) 2).setCellValue("STARTPCT1");
			rowhead1.createCell((short) 3).setCellValue("STARTPCT2");
			rowhead1.createCell((short) 4).setCellValue("STARTPCT3");
			rowhead1.createCell((short) 5).setCellValue("STARTPCT4");
			rowhead1.createCell((short) 6).setCellValue("STARTPCT5");
			rowhead1.createCell((short) 7).setCellValue("STARTPCT6");
			rowhead1.createCell((short) 8).setCellValue("DMDCAL");
			rowhead1.createCell((short) 9).setCellValue("STARTDATE");
			rowhead1.createCell((short) 10).setCellValue("DESCR");
			rowhead1.createCell((short) 11).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 12).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 13).setCellValue("LOC");
			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("DDEPROFILEID"));
				row.createCell((short) 2).setCellValue(resSet.getString("STARTPCT1"));
				row.createCell((short) 3).setCellValue(resSet.getString("STARTPCT2"));
				row.createCell((short) 4).setCellValue(resSet.getString("STARTPCT3"));
				row.createCell((short) 5).setCellValue(resSet.getString("STARTPCT4"));
				row.createCell((short) 6).setCellValue(resSet.getString("STARTPCT5"));
				row.createCell((short) 7).setCellValue(resSet.getString("STARTPCT6"));
				row.createCell((short) 8).setCellValue(resSet.getString("DMDCAL"));
				row.createCell((short) 9).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 10).setCellValue(resSet.getString("DESCR"));	
				row.createCell((short) 11).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 12).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 13).setCellValue(resSet.getString("LOC"));
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void DFU_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFU");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
			//rowhead1.createCell((short) 0).setCellValue("HISTSTART");
			rowhead1.createCell((short) 0).setCellValue("U_NP_ID");
			rowhead1.createCell((short) 1).setCellValue("U_USER_HISTSTART");
			rowhead1.createCell((short) 2).setCellValue("SEASONPROFILE");
			rowhead1.createCell((short) 3).setCellValue("U_PUBLISH_SW");
			rowhead1.createCell((short) 4).setCellValue("TOTFCSTLOCK");
			rowhead1.createCell((short) 5).setCellValue("LOCKDUR");
			rowhead1.createCell((short) 6).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 7).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 8).setCellValue("LOC");
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				//rowhead1.createCell((short) 0).setCellValue("HISTSTART");
				rowhead1.createCell((short) 0).setCellValue("U_NP_ID");
				rowhead1.createCell((short) 1).setCellValue("U_USER_HISTSTART");
				rowhead1.createCell((short) 2).setCellValue("SEASONPROFILE");
				rowhead1.createCell((short) 3).setCellValue("U_PUBLISH_SW");
				rowhead1.createCell((short) 4).setCellValue("TOTFCSTLOCK");
				rowhead1.createCell((short) 5).setCellValue("LOCKDUR");
				rowhead1.createCell((short) 6).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 7).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 8).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					//row.createCell((short) 0).setCellValue(resSet.getString("HISTSTART"));
					row.createCell((short) 0).setCellValue(resSet.getString("U_NP_ID"));
					row.createCell((short) 1).setCellValue(resSet.getString("U_USER_HISTSTART"));
					row.createCell((short) 2).setCellValue(resSet.getString("SEASONPROFILE"));
					row.createCell((short) 3).setCellValue(resSet.getString("U_PUBLISH_SW"));
					row.createCell((short) 4).setCellValue(resSet.getString("TOTFCSTLOCK"));
					row.createCell((short) 5).setCellValue(resSet.getString("LOCKDUR"));
					row.createCell((short) 6).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 7).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 8).setCellValue(resSet.getString("LOC"));
								
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				//row.createCell((short) 0).setCellValue(resSet.getString("HISTSTART"));
				row.createCell((short) 0).setCellValue(resSet.getString("U_NP_ID"));
				row.createCell((short) 1).setCellValue(resSet.getString("U_USER_HISTSTART"));
				row.createCell((short) 2).setCellValue(resSet.getString("SEASONPROFILE"));
				row.createCell((short) 3).setCellValue(resSet.getString("U_PUBLISH_SW"));
				row.createCell((short) 4).setCellValue(resSet.getString("TOTFCSTLOCK"));
				row.createCell((short) 5).setCellValue(resSet.getString("LOCKDUR"));
				row.createCell((short) 6).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 7).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 8).setCellValue(resSet.getString("LOC"));			
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void DFU_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFU");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			//rowhead1.createCell((short) 0).setCellValue("HISTSTART");
			rowhead1.createCell((short) 0).setCellValue("U_NP_ID");
			rowhead1.createCell((short) 1).setCellValue("U_USER_HISTSTART");
			rowhead1.createCell((short) 2).setCellValue("SEASONPROFILE");
			rowhead1.createCell((short) 3).setCellValue("U_PUBLISH_SW");
			rowhead1.createCell((short) 4).setCellValue("TOTFCSTLOCK");
			rowhead1.createCell((short) 5).setCellValue("LOCKDUR");
			rowhead1.createCell((short) 6).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 7).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 8).setCellValue("LOC");
			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				//row.createCell((short) 0).setCellValue(resSet.getString("HISTSTART"));
				row.createCell((short) 0).setCellValue(resSet.getString("U_NP_ID"));
				row.createCell((short) 1).setCellValue(resSet.getString("U_USER_HISTSTART"));
				row.createCell((short) 2).setCellValue(resSet.getString("SEASONPROFILE"));
				row.createCell((short) 3).setCellValue(resSet.getString("U_PUBLISH_SW"));
				row.createCell((short) 4).setCellValue(resSet.getString("TOTFCSTLOCK"));
				row.createCell((short) 5).setCellValue(resSet.getString("LOCKDUR"));
				row.createCell((short) 6).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 7).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 8).setCellValue(resSet.getString("LOC"));			
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void Eventmap_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFUMOVINGEVENTMAP");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
			rowhead1.createCell((short) 0).setCellValue("MOVINGEVENT");
			rowhead1.createCell((short) 1).setCellValue("OVERLAPFACTOR ");
			rowhead1.createCell((short) 2).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 3).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 4).setCellValue("LOC");
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				rowhead1.createCell((short) 0).setCellValue("MOVINGEVENT");
				rowhead1.createCell((short) 1).setCellValue("OVERLAPFACTOR");
				rowhead1.createCell((short) 2).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 3).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 4).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("MOVINGEVENT"));
					row.createCell((short) 1).setCellValue(resSet.getString("OVERLAPFACTOR"));
					row.createCell((short) 2).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 3).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 4).setCellValue(resSet.getString("LOC"));								
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				
				row.createCell((short) 0).setCellValue(resSet.getString("MOVINGEVENT"));
				row.createCell((short) 1).setCellValue(resSet.getString("OVERLAPFACTOR"));
				row.createCell((short) 2).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 3).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 4).setCellValue(resSet.getString("LOC"));						
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void Eventmap_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFUMOVINGEVENTMAP");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("MOVINGEVENT");
			rowhead1.createCell((short) 1).setCellValue("OVERLAPFACTOR");
			rowhead1.createCell((short) 2).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 3).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 4).setCellValue("LOC");			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MOVINGEVENT"));
				row.createCell((short) 1).setCellValue(resSet.getString("OVERLAPFACTOR"));
				row.createCell((short) 2).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 3).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 4).setCellValue(resSet.getString("LOC"));							
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		
		public void DFUDDE_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFUDDEMAP");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
				
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("DDEPROFILEID");
			rowhead1.createCell((short) 2).setCellValue("OPTIMPCT1");
			rowhead1.createCell((short) 3).setCellValue("OPTIMPCT2");
			rowhead1.createCell((short) 4).setCellValue("OPTIMPCT3");
			rowhead1.createCell((short) 5).setCellValue("OPTIMPCT4");
			rowhead1.createCell((short) 6).setCellValue("OPTIMPCT5");
			rowhead1.createCell((short) 7).setCellValue("OPTIMPCT6");
			rowhead1.createCell((short) 8).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 9).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 10).setCellValue("LOC");	
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				rowhead1.createCell((short) 0).setCellValue("MODEL");
				rowhead1.createCell((short) 1).setCellValue("DDEPROFILEID");
				rowhead1.createCell((short) 2).setCellValue("OPTIMPCT1");
				rowhead1.createCell((short) 3).setCellValue("OPTIMPCT2");
				rowhead1.createCell((short) 4).setCellValue("OPTIMPCT3");
				rowhead1.createCell((short) 5).setCellValue("OPTIMPCT4");
				rowhead1.createCell((short) 6).setCellValue("OPTIMPCT5");
				rowhead1.createCell((short) 7).setCellValue("OPTIMPCT6");
				rowhead1.createCell((short) 8).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 9).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 10).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
					row.createCell((short) 1).setCellValue(resSet.getString("DDEPROFILEID"));
					row.createCell((short) 2).setCellValue(resSet.getString("OPTIMPCT1"));
					row.createCell((short) 3).setCellValue(resSet.getString("OPTIMPCT2"));
					row.createCell((short) 4).setCellValue(resSet.getString("OPTIMPCT3"));
					row.createCell((short) 5).setCellValue(resSet.getString("OPTIMPCT4"));
					row.createCell((short) 6).setCellValue(resSet.getString("OPTIMPCT5"));
					row.createCell((short) 7).setCellValue(resSet.getString("OPTIMPCT6"));
					row.createCell((short) 8).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 9).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 10).setCellValue(resSet.getString("LOC"));
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("DDEPROFILEID"));
				row.createCell((short) 2).setCellValue(resSet.getString("OPTIMPCT1"));
				row.createCell((short) 3).setCellValue(resSet.getString("OPTIMPCT2"));
				row.createCell((short) 4).setCellValue(resSet.getString("OPTIMPCT3"));
				row.createCell((short) 5).setCellValue(resSet.getString("OPTIMPCT4"));
				row.createCell((short) 6).setCellValue(resSet.getString("OPTIMPCT5"));
				row.createCell((short) 7).setCellValue(resSet.getString("OPTIMPCT6"));
				row.createCell((short) 8).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 9).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 10).setCellValue(resSet.getString("LOC"));						
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void DFUDDE_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFUDDEMAP");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("DDEPROFILEID");
			rowhead1.createCell((short) 2).setCellValue("OPTIMPCT1");
			rowhead1.createCell((short) 3).setCellValue("OPTIMPCT2");
			rowhead1.createCell((short) 4).setCellValue("OPTIMPCT3");
			rowhead1.createCell((short) 5).setCellValue("OPTIMPCT4");
			rowhead1.createCell((short) 6).setCellValue("OPTIMPCT5");
			rowhead1.createCell((short) 7).setCellValue("OPTIMPCT6");
			rowhead1.createCell((short) 8).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 9).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 10).setCellValue("LOC");			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("DDEPROFILEID"));
				row.createCell((short) 2).setCellValue(resSet.getString("OPTIMPCT1"));
				row.createCell((short) 3).setCellValue(resSet.getString("OPTIMPCT2"));
				row.createCell((short) 4).setCellValue(resSet.getString("OPTIMPCT3"));
				row.createCell((short) 5).setCellValue(resSet.getString("OPTIMPCT4"));
				row.createCell((short) 6).setCellValue(resSet.getString("OPTIMPCT5"));
				row.createCell((short) 7).setCellValue(resSet.getString("OPTIMPCT6"));
				row.createCell((short) 8).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 9).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 10).setCellValue(resSet.getString("LOC"));							
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void DFUEFF_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFUEFFPRICE");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
			rowhead1.createCell((short) 0).setCellValue("STARTDATE");
			rowhead1.createCell((short) 1).setCellValue("EFFPRICE");
			rowhead1.createCell((short) 2).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 3).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 4).setCellValue("LOC");
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				rowhead1.createCell((short) 0).setCellValue("STARTDATE");
				rowhead1.createCell((short) 1).setCellValue("EFFPRICE");
				rowhead1.createCell((short) 2).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 3).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 4).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("STARTDATE"));
					row.createCell((short) 1).setCellValue(resSet.getString("EFFPRICE"));
					row.createCell((short) 2).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 3).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 4).setCellValue(resSet.getString("LOC"));								
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				
				row.createCell((short) 0).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 1).setCellValue(resSet.getString("EFFPRICE"));
				row.createCell((short) 2).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 3).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 4).setCellValue(resSet.getString("LOC"));						
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void DFUEFF_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("DFUEFFPRICE");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("STARTDATE");
			rowhead1.createCell((short) 1).setCellValue("EFFPRICE");
			rowhead1.createCell((short) 2).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 3).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 4).setCellValue("LOC");			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 1).setCellValue(resSet.getString("EFFPRICE"));
				row.createCell((short) 2).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 3).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 4).setCellValue(resSet.getString("LOC"));							
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void FCST_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("FCSTnew");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
				
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("STARTDATE");
			rowhead1.createCell((short) 2).setCellValue("DUR");
			rowhead1.createCell((short) 3).setCellValue("TYPE");
			rowhead1.createCell((short) 4).setCellValue("FCSTID");
			rowhead1.createCell((short) 5).setCellValue("QTY");
			rowhead1.createCell((short) 6).setCellValue("LEWMEANQTY");
			rowhead1.createCell((short) 7).setCellValue("MARKETMGRVERSIONID");
			rowhead1.createCell((short) 8).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 9).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 10).setCellValue("LOC");
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				rowhead1.createCell((short) 0).setCellValue("MODEL");
				rowhead1.createCell((short) 1).setCellValue("STARTDATE");
				rowhead1.createCell((short) 2).setCellValue("DUR");
				rowhead1.createCell((short) 3).setCellValue("TYPE");
				rowhead1.createCell((short) 4).setCellValue("FCSTID");
				rowhead1.createCell((short) 5).setCellValue("QTY");
				rowhead1.createCell((short) 6).setCellValue("LEWMEANQTY");
				rowhead1.createCell((short) 7).setCellValue("MARKETMGRVERSIONID");
				rowhead1.createCell((short) 8).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 9).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 10).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
					row.createCell((short) 1).setCellValue(resSet.getString("STARTDATE"));
					row.createCell((short) 2).setCellValue(resSet.getString("DUR"));
					row.createCell((short) 3).setCellValue(resSet.getString("TYPE"));
					row.createCell((short) 4).setCellValue(resSet.getString("FCSTID"));
					row.createCell((short) 5).setCellValue(resSet.getString("QTY"));
					row.createCell((short) 6).setCellValue(resSet.getString("LEWMEANQTY"));
					row.createCell((short) 7).setCellValue(resSet.getString("MARKETMGRVERSIONID"));
					row.createCell((short) 8).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 9).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 10).setCellValue(resSet.getString("LOC"));
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 2).setCellValue(resSet.getString("DUR"));
				row.createCell((short) 3).setCellValue(resSet.getString("TYPE"));
				row.createCell((short) 4).setCellValue(resSet.getString("FCSTID"));
				row.createCell((short) 5).setCellValue(resSet.getString("QTY"));
				row.createCell((short) 6).setCellValue(resSet.getString("LEWMEANQTY"));
				row.createCell((short) 7).setCellValue(resSet.getString("MARKETMGRVERSIONID"));
				row.createCell((short) 8).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 9).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 10).setCellValue(resSet.getString("LOC"));						
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void FCST_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("FCSTnew");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("STARTDATE");
			rowhead1.createCell((short) 2).setCellValue("DUR");
			rowhead1.createCell((short) 3).setCellValue("TYPE");
			rowhead1.createCell((short) 4).setCellValue("FCSTID");
			rowhead1.createCell((short) 5).setCellValue("QTY");
			rowhead1.createCell((short) 6).setCellValue("LEWMEANQTY");
			rowhead1.createCell((short) 7).setCellValue("MARKETMGRVERSIONID");
			rowhead1.createCell((short) 8).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 9).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 10).setCellValue("LOC");			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 2).setCellValue(resSet.getString("DUR"));
				row.createCell((short) 3).setCellValue(resSet.getString("TYPE"));
				row.createCell((short) 4).setCellValue(resSet.getString("FCSTID"));
				row.createCell((short) 5).setCellValue(resSet.getString("QTY"));
				row.createCell((short) 6).setCellValue(resSet.getString("LEWMEANQTY"));
				row.createCell((short) 7).setCellValue(resSet.getString("MARKETMGRVERSIONID"));
				row.createCell((short) 8).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 9).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 10).setCellValue(resSet.getString("LOC"));							
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void LEWAND_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("lewandowskiparam");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
			rowhead1.createCell((short) 0).setCellValue("INITIALLINEARTREND");
			rowhead1.createCell((short) 1).setCellValue("INITIALQUADTREND");
			rowhead1.createCell((short) 2).setCellValue("MEANVALUEDYNAMIC");
			rowhead1.createCell((short) 3).setCellValue("MEANVALUEMAX");
			rowhead1.createCell((short) 4).setCellValue("SEASONALITYIMPACT");
			rowhead1.createCell((short) 5).setCellValue("TRENDCOMBINATION");
			rowhead1.createCell((short) 6).setCellValue("HYBRIDFACTOR");
			rowhead1.createCell((short) 7).setCellValue("TRACKINGSIGNALAWS");
			rowhead1.createCell((short) 8).setCellValue("STABILITYRATENF");
			rowhead1.createCell((short) 9).setCellValue("SMOOTHEDMAD");
			rowhead1.createCell((short) 10).setCellValue("HISTDEPENDENCY");
			rowhead1.createCell((short) 11).setCellValue("DDEIMPACT");
			rowhead1.createCell((short) 12).setCellValue("LINFACTOROPT");
			rowhead1.createCell((short) 13).setCellValue("LIFECYCLEFACTOR");
			rowhead1.createCell((short) 14).setCellValue("LINEXTFACTOR");
			rowhead1.createCell((short) 15).setCellValue("NONLINEXTFACTOR");
			rowhead1.createCell((short) 16).setCellValue("LIFECYCLESTARTDATE");
			rowhead1.createCell((short) 17).setCellValue("LINFACTORIMPACT");
			rowhead1.createCell((short) 18).setCellValue("NONLINFACTORAMP");
			rowhead1.createCell((short) 19).setCellValue("NONLINFACTRESPONSE");
			rowhead1.createCell((short) 20).setCellValue("SMOOTHPROFILESW");
			rowhead1.createCell((short) 21).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 22).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 23).setCellValue("LOC");
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				rowhead1.createCell((short) 0).setCellValue("INITIALLINEARTREND");
				rowhead1.createCell((short) 1).setCellValue("INITIALQUADTREND");
				rowhead1.createCell((short) 2).setCellValue("MEANVALUEDYNAMIC");
				rowhead1.createCell((short) 3).setCellValue("MEANVALUEMAX");
				rowhead1.createCell((short) 4).setCellValue("SEASONALITYIMPACT");
				rowhead1.createCell((short) 5).setCellValue("TRENDCOMBINATION");
				rowhead1.createCell((short) 6).setCellValue("HYBRIDFACTOR");
				rowhead1.createCell((short) 7).setCellValue("TRACKINGSIGNALAWS");
				rowhead1.createCell((short) 8).setCellValue("STABILITYRATENF");
				rowhead1.createCell((short) 9).setCellValue("SMOOTHEDMAD");
				rowhead1.createCell((short) 10).setCellValue("HISTDEPENDENCY");
				rowhead1.createCell((short) 11).setCellValue("DDEIMPACT");
				rowhead1.createCell((short) 12).setCellValue("LINFACTOROPT");
				rowhead1.createCell((short) 13).setCellValue("LIFECYCLEFACTOR");
				rowhead1.createCell((short) 14).setCellValue("LINEXTFACTOR");
				rowhead1.createCell((short) 15).setCellValue("NONLINEXTFACTOR");
				rowhead1.createCell((short) 16).setCellValue("LIFECYCLESTARTDATE");
				rowhead1.createCell((short) 17).setCellValue("LINFACTORIMPACT");
				rowhead1.createCell((short) 18).setCellValue("NONLINFACTORAMP");
				rowhead1.createCell((short) 19).setCellValue("NONLINFACTRESPONSE");
				rowhead1.createCell((short) 20).setCellValue("SMOOTHPROFILESW");
				rowhead1.createCell((short) 21).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 22).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 23).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("INITIALLINEARTREND"));
					row.createCell((short) 1).setCellValue(resSet.getString("INITIALQUADTREND"));
					row.createCell((short) 2).setCellValue(resSet.getString("MEANVALUEDYNAMIC"));
					row.createCell((short) 3).setCellValue(resSet.getString("MEANVALUEMAX"));
					row.createCell((short) 4).setCellValue(resSet.getString("SEASONALITYIMPACT"));
					row.createCell((short) 5).setCellValue(resSet.getString("TRENDCOMBINATION"));
					row.createCell((short) 6).setCellValue(resSet.getString("HYBRIDFACTOR"));
					row.createCell((short) 7).setCellValue(resSet.getString("TRACKINGSIGNALAWS"));
					row.createCell((short) 8).setCellValue(resSet.getString("STABILITYRATENF"));
					row.createCell((short) 9).setCellValue(resSet.getString("SMOOTHEDMAD"));
					row.createCell((short) 10).setCellValue(resSet.getString("HISTDEPENDENCY"));
					row.createCell((short) 11).setCellValue(resSet.getString("DDEIMPACT"));
					row.createCell((short) 12).setCellValue(resSet.getString("LINFACTOROPT"));
					row.createCell((short) 13).setCellValue(resSet.getString("LIFECYCLEFACTOR"));
					row.createCell((short) 14).setCellValue(resSet.getString("LINEXTFACTOR"));
					row.createCell((short) 15).setCellValue(resSet.getString("NONLINEXTFACTOR"));
					row.createCell((short) 16).setCellValue(resSet.getString("LIFECYCLESTARTDATE"));
					row.createCell((short) 17).setCellValue(resSet.getString("LINFACTORIMPACT"));
					row.createCell((short) 18).setCellValue(resSet.getString("NONLINFACTORAMP"));
					row.createCell((short) 19).setCellValue(resSet.getString("NONLINFACTRESPONSE"));
					row.createCell((short) 20).setCellValue(resSet.getString("SMOOTHPROFILESW"));
					row.createCell((short) 21).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 22).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 23).setCellValue(resSet.getString("LOC"));
					
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("INITIALLINEARTREND"));
				row.createCell((short) 1).setCellValue(resSet.getString("INITIALQUADTREND"));
				row.createCell((short) 2).setCellValue(resSet.getString("MEANVALUEDYNAMIC"));
				row.createCell((short) 3).setCellValue(resSet.getString("MEANVALUEMAX"));
				row.createCell((short) 4).setCellValue(resSet.getString("SEASONALITYIMPACT"));
				row.createCell((short) 5).setCellValue(resSet.getString("TRENDCOMBINATION"));
				row.createCell((short) 6).setCellValue(resSet.getString("HYBRIDFACTOR"));
				row.createCell((short) 7).setCellValue(resSet.getString("TRACKINGSIGNALAWS"));
				row.createCell((short) 8).setCellValue(resSet.getString("STABILITYRATENF"));
				row.createCell((short) 9).setCellValue(resSet.getString("SMOOTHEDMAD"));
				row.createCell((short) 10).setCellValue(resSet.getString("HISTDEPENDENCY"));
				row.createCell((short) 11).setCellValue(resSet.getString("DDEIMPACT"));
				row.createCell((short) 12).setCellValue(resSet.getString("LINFACTOROPT"));
				row.createCell((short) 13).setCellValue(resSet.getString("LIFECYCLEFACTOR"));
				row.createCell((short) 14).setCellValue(resSet.getString("LINEXTFACTOR"));
				row.createCell((short) 15).setCellValue(resSet.getString("NONLINEXTFACTOR"));
				row.createCell((short) 16).setCellValue(resSet.getString("LIFECYCLESTARTDATE"));
				row.createCell((short) 17).setCellValue(resSet.getString("LINFACTORIMPACT"));
				row.createCell((short) 18).setCellValue(resSet.getString("NONLINFACTORAMP"));
				row.createCell((short) 19).setCellValue(resSet.getString("NONLINFACTRESPONSE"));
				row.createCell((short) 20).setCellValue(resSet.getString("SMOOTHPROFILESW"));
				row.createCell((short) 21).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 22).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 23).setCellValue(resSet.getString("LOC"));						
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void LEWAND_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("lewandowskiparam");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("INITIALLINEARTREND");
			rowhead1.createCell((short) 1).setCellValue("INITIALQUADTREND");
			rowhead1.createCell((short) 2).setCellValue("MEANVALUEDYNAMIC");
			rowhead1.createCell((short) 3).setCellValue("MEANVALUEMAX");
			rowhead1.createCell((short) 4).setCellValue("SEASONALITYIMPACT");
			rowhead1.createCell((short) 5).setCellValue("TRENDCOMBINATION");
			rowhead1.createCell((short) 6).setCellValue("HYBRIDFACTOR");
			rowhead1.createCell((short) 7).setCellValue("TRACKINGSIGNALAWS");
			rowhead1.createCell((short) 8).setCellValue("STABILITYRATENF");
			rowhead1.createCell((short) 9).setCellValue("SMOOTHEDMAD");
			rowhead1.createCell((short) 10).setCellValue("HISTDEPENDENCY");
			rowhead1.createCell((short) 11).setCellValue("DDEIMPACT");
			rowhead1.createCell((short) 12).setCellValue("LINFACTOROPT");
			rowhead1.createCell((short) 13).setCellValue("LIFECYCLEFACTOR");
			rowhead1.createCell((short) 14).setCellValue("LINEXTFACTOR");
			rowhead1.createCell((short) 15).setCellValue("NONLINEXTFACTOR");
			rowhead1.createCell((short) 16).setCellValue("LIFECYCLESTARTDATE");
			rowhead1.createCell((short) 17).setCellValue("LINFACTORIMPACT");
			rowhead1.createCell((short) 18).setCellValue("NONLINFACTORAMP");
			rowhead1.createCell((short) 19).setCellValue("NONLINFACTRESPONSE");
			rowhead1.createCell((short) 20).setCellValue("SMOOTHPROFILESW");
			rowhead1.createCell((short) 21).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 22).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 23).setCellValue("LOC");			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("INITIALLINEARTREND"));
				row.createCell((short) 1).setCellValue(resSet.getString("INITIALQUADTREND"));
				row.createCell((short) 2).setCellValue(resSet.getString("MEANVALUEDYNAMIC"));
				row.createCell((short) 3).setCellValue(resSet.getString("MEANVALUEMAX"));
				row.createCell((short) 4).setCellValue(resSet.getString("SEASONALITYIMPACT"));
				row.createCell((short) 5).setCellValue(resSet.getString("TRENDCOMBINATION"));
				row.createCell((short) 6).setCellValue(resSet.getString("HYBRIDFACTOR"));
				row.createCell((short) 7).setCellValue(resSet.getString("TRACKINGSIGNALAWS"));
				row.createCell((short) 8).setCellValue(resSet.getString("STABILITYRATENF"));
				row.createCell((short) 9).setCellValue(resSet.getString("SMOOTHEDMAD"));
				row.createCell((short) 10).setCellValue(resSet.getString("HISTDEPENDENCY"));
				row.createCell((short) 11).setCellValue(resSet.getString("DDEIMPACT"));
				row.createCell((short) 12).setCellValue(resSet.getString("LINFACTOROPT"));
				row.createCell((short) 13).setCellValue(resSet.getString("LIFECYCLEFACTOR"));
				row.createCell((short) 14).setCellValue(resSet.getString("LINEXTFACTOR"));
				row.createCell((short) 15).setCellValue(resSet.getString("NONLINEXTFACTOR"));
				row.createCell((short) 16).setCellValue(resSet.getString("LIFECYCLESTARTDATE"));
				row.createCell((short) 17).setCellValue(resSet.getString("LINFACTORIMPACT"));
				row.createCell((short) 18).setCellValue(resSet.getString("NONLINFACTORAMP"));
				row.createCell((short) 19).setCellValue(resSet.getString("NONLINFACTRESPONSE"));
				row.createCell((short) 20).setCellValue(resSet.getString("SMOOTHPROFILESW"));
				row.createCell((short) 21).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 22).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 23).setCellValue(resSet.getString("LOC"));							
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void Meanvalue_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("meanvalueadj");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("STARTDATE");
			rowhead1.createCell((short) 2).setCellValue("DESCR");
			rowhead1.createCell((short) 3).setCellValue("MODTYPE");
			rowhead1.createCell((short) 4).setCellValue("ADJRATE");
			rowhead1.createCell((short) 5).setCellValue("FIXINFUTURESW");
			rowhead1.createCell((short) 6).setCellValue("ADJVAL");
			rowhead1.createCell((short) 7).setCellValue("DMDCAL");
			rowhead1.createCell((short) 8).setCellValue("FIXUPTONUMPERIODSSW");
			rowhead1.createCell((short) 9).setCellValue("NUMPERIODS");
			rowhead1.createCell((short) 10).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 11).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 12).setCellValue("LOC");
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				rowhead1.createCell((short) 0).setCellValue("MODEL");
				rowhead1.createCell((short) 1).setCellValue("STARTDATE");
				rowhead1.createCell((short) 2).setCellValue("DESCR");
				rowhead1.createCell((short) 3).setCellValue("MODTYPE");
				rowhead1.createCell((short) 4).setCellValue("ADJRATE");
				rowhead1.createCell((short) 5).setCellValue("FIXINFUTURESW");
				rowhead1.createCell((short) 6).setCellValue("ADJVAL");
				rowhead1.createCell((short) 7).setCellValue("DMDCAL");
				rowhead1.createCell((short) 8).setCellValue("FIXUPTONUMPERIODSSW");
				rowhead1.createCell((short) 9).setCellValue("NUMPERIODS");
				rowhead1.createCell((short) 10).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 11).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 12).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
					row.createCell((short) 1).setCellValue(resSet.getString("STARTDATE"));
					row.createCell((short) 2).setCellValue(resSet.getString("DESCR"));
					row.createCell((short) 3).setCellValue(resSet.getString("MODTYPE"));
					row.createCell((short) 4).setCellValue(resSet.getString("ADJRATE"));
					row.createCell((short) 5).setCellValue(resSet.getString("FIXINFUTURESW"));
					row.createCell((short) 6).setCellValue(resSet.getString("ADJVAL"));
					row.createCell((short) 7).setCellValue(resSet.getString("DMDCAL"));
					row.createCell((short) 8).setCellValue(resSet.getString("FIXUPTONUMPERIODSSW"));
					row.createCell((short) 9).setCellValue(resSet.getString("NUMPERIODS"));
					row.createCell((short) 10).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 11).setCellValue(resSet.getString("DMDGROUP"));
					row.createCell((short) 12).setCellValue(resSet.getString("LOC"));
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 2).setCellValue(resSet.getString("DESCR"));
				row.createCell((short) 3).setCellValue(resSet.getString("MODTYPE"));
				row.createCell((short) 4).setCellValue(resSet.getString("ADJRATE"));
				row.createCell((short) 5).setCellValue(resSet.getString("FIXINFUTURESW"));
				row.createCell((short) 6).setCellValue(resSet.getString("ADJVAL"));
				row.createCell((short) 7).setCellValue(resSet.getString("DMDCAL"));
				row.createCell((short) 8).setCellValue(resSet.getString("FIXUPTONUMPERIODSSW"));
				row.createCell((short) 9).setCellValue(resSet.getString("NUMPERIODS"));
				row.createCell((short) 10).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 11).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 12).setCellValue(resSet.getString("LOC"));						
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void Meanvalue_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("meanvalueadj");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("MODEL");
			rowhead1.createCell((short) 1).setCellValue("STARTDATE");
			rowhead1.createCell((short) 2).setCellValue("DESCR");
			rowhead1.createCell((short) 3).setCellValue("MODTYPE");
			rowhead1.createCell((short) 4).setCellValue("ADJRATE");
			rowhead1.createCell((short) 5).setCellValue("FIXINFUTURESW");
			rowhead1.createCell((short) 6).setCellValue("ADJVAL");
			rowhead1.createCell((short) 7).setCellValue("DMDCAL");
			rowhead1.createCell((short) 8).setCellValue("FIXUPTONUMPERIODSSW");
			rowhead1.createCell((short) 9).setCellValue("NUMPERIODS");
			rowhead1.createCell((short) 10).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 11).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 12).setCellValue("LOC");			
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("MODEL"));
				row.createCell((short) 1).setCellValue(resSet.getString("STARTDATE"));
				row.createCell((short) 2).setCellValue(resSet.getString("DESCR"));
				row.createCell((short) 3).setCellValue(resSet.getString("MODTYPE"));
				row.createCell((short) 4).setCellValue(resSet.getString("ADJRATE"));
				row.createCell((short) 5).setCellValue(resSet.getString("FIXINFUTURESW"));
				row.createCell((short) 6).setCellValue(resSet.getString("ADJVAL"));
				row.createCell((short) 7).setCellValue(resSet.getString("DMDCAL"));
				row.createCell((short) 8).setCellValue(resSet.getString("FIXUPTONUMPERIODSSW"));
				row.createCell((short) 9).setCellValue(resSet.getString("NUMPERIODS"));
				row.createCell((short) 10).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 11).setCellValue(resSet.getString("DMDGROUP"));
				row.createCell((short) 12).setCellValue(resSet.getString("LOC"));							
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void targetdfu_Validation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("targetdfumap");
			int rowmax=sheet1.getLastRowNum();
			System.out.println(rowmax);
			int newrowmax = rowmax+5;
			if(rowmax==0) {
			Row rowhead1 = sheet1.createRow((short) 0);
			rowhead1.createCell((short) 0).setCellValue("model");
			rowhead1.createCell((short) 1).setCellValue("target");
			rowhead1.createCell((short) 2).setCellValue("qty");
			rowhead1.createCell((short) 3).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 4).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 5).setCellValue("LOC");
			}
			else {
				
				Row rowhead1 = sheet1.createRow((short) newrowmax);
				rowhead1.createCell((short) 0).setCellValue("model");
				rowhead1.createCell((short) 1).setCellValue("target");
				rowhead1.createCell((short) 2).setCellValue("qty");
				rowhead1.createCell((short) 3).setCellValue("DMDUNIT");
				rowhead1.createCell((short) 4).setCellValue("DMDGROUP");
				rowhead1.createCell((short) 5).setCellValue("LOC");
			}
								
			if(rowmax==0) {
				int i=rowmax+1;
				while (resSet.next())
				{		
					Row row = sheet1.createRow((short) i++);
					row.createCell((short) 0).setCellValue(resSet.getString("model"));
					row.createCell((short) 1).setCellValue(resSet.getString("target"));
					row.createCell((short) 2).setCellValue(resSet.getString("qty"));
					row.createCell((short) 3).setCellValue(resSet.getString("DMDUNIT"));
					row.createCell((short) 4).setCellValue(resSet.getString("DMDGROUP"));	
					row.createCell((short) 5).setCellValue(resSet.getString("LOC"));	
					}
			}
			else {
			int i=newrowmax+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				
				row.createCell((short) 0).setCellValue(resSet.getString("model"));
				row.createCell((short) 1).setCellValue(resSet.getString("target"));
				row.createCell((short) 2).setCellValue(resSet.getString("qty"));
				row.createCell((short) 3).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 4).setCellValue(resSet.getString("DMDGROUP"));	
				row.createCell((short) 5).setCellValue(resSet.getString("LOC"));						
				}
			}	
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		
		public void targetdfu_SUPSValidation(String Query) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet("targetdfumap");
			int rowmax=sheet1.getLastRowNum();
			System.out.println("rowmax= " +rowmax );
			int j= rowmax+4;
			Row rowhead1 = sheet1.createRow((short) j);
			rowhead1.createCell((short) 0).setCellValue("model");
			rowhead1.createCell((short) 1).setCellValue("target");
			rowhead1.createCell((short) 2).setCellValue("qty");
			rowhead1.createCell((short) 3).setCellValue("DMDUNIT");
			rowhead1.createCell((short) 4).setCellValue("DMDGROUP");
			rowhead1.createCell((short) 5).setCellValue("LOC");		
			int i=j+1;		
			while (resSet.next())
			{		
				Row row = sheet1.createRow((short) i++);
				row.createCell((short) 0).setCellValue(resSet.getString("model"));
				row.createCell((short) 1).setCellValue(resSet.getString("target"));
				row.createCell((short) 2).setCellValue(resSet.getString("qty"));
				row.createCell((short) 3).setCellValue(resSet.getString("DMDUNIT"));
				row.createCell((short) 4).setCellValue(resSet.getString("DMDGROUP"));	
				row.createCell((short) 5).setCellValue(resSet.getString("LOC"));							
				}
					
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}
		public void minusquery1(String Query, String Sheet1) throws SQLException, IOException
		{
			dbopenSnap();
			Connection con=DriverManager.getConnection(connectionURL,userName,pass);
			Statement stmt=con.createStatement();
			report.log(Query);
			ResultSet resSet = stmt.executeQuery(Query);
			Sheet sheet1 = workbook.getSheet(Sheet1);
			int rowmax=sheet1.getLastRowNum();
			int Cmprst = rowmax+2;
			if (resSet.next()) {
				
				Row row = sheet1.createRow((short) Cmprst);
				row.createCell((short) 0).setCellValue("Mismatched Snapshots, Hence Failed");
			}
			else {
				Row row = sheet1.createRow((short) Cmprst);
				row.createCell((short) 0).setCellValue("Snapshots Match, Hence Passed");
				
			}
			
			FileOutputStream fileOut = new FileOutputStream(snapShotFilePath);
			workbook.write(fileOut);
			con.close();
		}

		
		
}
