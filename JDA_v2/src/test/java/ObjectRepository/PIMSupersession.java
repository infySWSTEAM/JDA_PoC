package ObjectRepository;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import Functions.ExcelFile;

public class PIMSupersession {
	
	WebDriver driver = null;
	ExcelFile exfile = new ExcelFile();
	WebDriverWait wait = null;
	
	String ItemNumber;
	String ReplacedId;
	String UltimateParent;

    By item_number = By.xpath("//*[@id='PopupAppDiv']/form/table[2]/tbody/tr[2]/td[3]/input");
    By replaced_id = By.xpath("//*[@id='PopupAppDiv']/form/table[2]/tbody/tr[4]/td[3]/input");
    By ult_parent =By.xpath("//*[@id='PopupAppDiv']/form/table[2]/tbody/tr[6]/td[3]/input");
    By Done =By.xpath("//span[@id='PromptScreenPopup_doneBtnOn']/div/a/div/div[@class='j-button-fill']/span[text()='Done']");
    By Dropdown = By.xpath("//*[@title='UDT_PIM_ITEM_SUPERSESSION']");
    By Clear =By.xpath("//*[@id='PromptScreenPopup_clearOn']/div/div/span");
    
    public PIMSupersession(WebDriver driver){

        this.driver = driver;

    }
    
    public void dropdownclick() {
    	driver.findElement(Dropdown).click();
    }
    public void setItemNumber(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	ItemNumber = exfile.readExcel(filePa, fileNa, SheetNa, row, 2);
        driver.findElement(item_number).sendKeys(ItemNumber);

    }
    
    public void setReplacedId(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	ReplacedId = exfile.readExcel(filePa, fileNa, SheetNa, row, 3);
        driver.findElement(replaced_id).sendKeys(ReplacedId);

    }
    
    public void setUltimateParent(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	UltimateParent = exfile.readExcel(filePa, fileNa, SheetNa, row, 0);
        driver.findElement(ult_parent).sendKeys(UltimateParent);

    }
    
    public void clickdone() {
    	driver.findElement(Done).click();
    }
    
    public void Clear() {
    	driver.findElement(Clear).click();
    }
    
}
