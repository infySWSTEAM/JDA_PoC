package ObjectRepository;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import Functions.ExcelFile;

public class Supersession {
	
	WebDriver driver = null;
	ExcelFile exfile = new ExcelFile();
	WebDriverWait wait = null;
	
	String InputItem;
	String InputLocation;
	By Dropdown = By.xpath("//*[@id='directory']//span[@title='Supersession']");
    By inputItem = By.xpath("//*[@id='PopupAppDiv']/form/table[2]/tbody/tr[2]/td[3]/input");
    By inputloc = By.xpath("//*[@name='!SRC!SUPERSESSION!LOC!E!Supersession-Prompt!P!2']");
    By Done =By.xpath("//span[@id='PromptScreenPopup_doneBtnOn']/div/a/div/div[@class='j-button-fill']/span[text()='Done']");
    By Clear =By.xpath("//*[@id='PromptScreenPopup_clearOn']/div/div/span");

            
    public Supersession(WebDriver driver){

        this.driver = driver;

    }
    
    public void dropdownclick() {
    	driver.findElement(Dropdown).click();
    }
    public void setInputItem(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	InputItem = exfile.readExcel(filePa, fileNa, SheetNa, row, 0);
        driver.findElement(inputItem).sendKeys(InputItem);

    }
    
    public void setInputLoc(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	InputLocation = exfile.readExcel(filePa, fileNa, SheetNa, row, 1);
        driver.findElement(inputloc).sendKeys(InputLocation);

    }
    
    public void clickdone() {
    	driver.findElement(Done).click();
    }
    
    public void Clear() {
    	driver.findElement(Clear).click();
    }
}
