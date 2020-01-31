package ObjectRepository;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import Functions.ExcelFile;

public class DemandWorkBench {
	
	WebDriver driver = null;
	ExcelFile exfile = new ExcelFile();
	WebDriverWait wait = null;
	
	String DMDUNIT;
	String DMDGROUP;
	String LOC;
    By inputDMDUNIT = By.xpath("//*[@id='PopupAppDiv']/form/table[2]/tbody/tr[5]/td[3]/input");
    By inputDMDGROUP = By.xpath("//*[@id='PopupAppDiv']/form/table[2]/tbody/tr[7]/td[3]/input");
    By inputLOC = By.xpath("//*[@id='PopupAppDiv']/form/table[2]/tbody/tr[11]/td[3]/input");
    By Done =By.xpath("//span[@id='PromptScreenPopup_doneBtnOn']/div/a/div/div[@class='j-button-fill']/span[text()='Done']");
    By Dropdown = By.xpath("//*[@title='Demand Workbench - New']");
    By Clear =By.xpath("//*[@id='PromptScreenPopup_clearOn']/div/div/span");
    
    public DemandWorkBench(WebDriver driver){

        this.driver = driver;

    }
    
    public void dropdownclick() {
    	driver.findElement(Dropdown).click();
    }
    
    public void setDMDUNIT(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	DMDUNIT = exfile.readExcel(filePa, fileNa, SheetNa, row, 0);
        driver.findElement(inputDMDUNIT).sendKeys(DMDUNIT);

    }
    
    public void setDMDGROUP(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	DMDGROUP = exfile.readExcel(filePa, fileNa, SheetNa, row, 2);
        driver.findElement(inputDMDGROUP).sendKeys(DMDGROUP);

    }
    
    public void setLOC(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	LOC = exfile.readExcel(filePa, fileNa, SheetNa, row, 1);
        driver.findElement(inputLOC).sendKeys(LOC);

    }
    
    public void clickdone() {
    	driver.findElement(Done).click();
    }
    
    public void Clear() {
    	driver.findElement(Clear).click();
    }
}
