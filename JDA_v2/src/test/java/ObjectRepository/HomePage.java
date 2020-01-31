package ObjectRepository;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;

import Functions.ExcelFile;

public class HomePage {

	WebDriver driver=null;
	ExcelFile exfile = new ExcelFile();
	String TableName;
	
	By uisearchbox = By.xpath("//*[@id=\'_directoryInput\']");
	 
    public HomePage(WebDriver driver){

        this.driver = driver;

    }
	 
	 public void setsearchname(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
		 TableName = exfile.readExcel(filePa, fileNa, SheetNa, row, 0);
		 driver.findElement(uisearchbox).sendKeys(TableName);
	     }

	 public void clr() {
		 driver.findElement(uisearchbox).clear();
			
	 }
}
