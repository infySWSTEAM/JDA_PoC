///************************************************************************************************************************
//		Author           : SGWS JDA Team 
//		Last Modified by : Nivedha Ravichandran
//		Last Modified on : 13-Feb-2020
//		Summary 		 : SQL Validations for SS Classification Rejection scenarios
//
//************************************************************************************************************************/

package ObjectRepository;

import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

import Functions.ExcelFile;

public class LoginPage {
	
	WebDriver driver = null;
	ExcelFile exfile = new ExcelFile();
	WebDriverWait wait = null;
	
	String strUserName;
	String strPassWord;

    By txtUserName = By.name("USER_NAME");
    By txtPassWord = By.name("PASSWORD");
    By loginbtn =By.className("loginInner");
    By useroption = By.xpath("//*[@id='shellUsername']");
    By Logout = By.xpath("//*[@id='shellLogoutAction']");
    
    public LoginPage(WebDriver driver){

        this.driver = driver;

    }
    public void useroption() {

    	driver.findElement(useroption).click();
    }
    
    public void JDAlogout() {

    	driver.findElement(Logout).click();
    }
            
    public void setUsername(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	strUserName = exfile.readExcel(filePa, fileNa, SheetNa, row, 0);
        driver.findElement(txtUserName).sendKeys(strUserName);

    }
    
    public void setPassword(String filePa, String fileNa,String SheetNa,int row,int col) throws IOException{
    	strPassWord = exfile.readExcel(filePa, fileNa, SheetNa, row, 1);
        driver.findElement(txtPassWord).sendKeys(strPassWord);

    }
    
    public void clickSubmit() {

    	driver.findElement(loginbtn).click();
    }
}
