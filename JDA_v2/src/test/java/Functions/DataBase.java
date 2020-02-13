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
	}
	
	public void newSkuCheck(String Query,String NewItem, int i) throws IOException, SQLException 
	{
		dbopen();
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
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
	public void dbSkuRejection(String Query) throws SQLException, IOException
	{
		dbopen();
		Connection con=DriverManager.getConnection(connectionURL,userName,pass);
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
}
