package testScript;

import org.testng.annotations.Test;
import org.testng.annotations.DataProvider;
import org.testng.annotations.BeforeTest;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import utilitiesXPO.ReadExcel;

public class TC001_LoginOTM {
  private static WebDriver driver;
  public static WebDriver getDriver() {
	  return driver;
  }
  @BeforeSuite
  public void LaunchOTM() {
	String baseUrl = "https://otmq.xpo.com";
	//System.setProperty("webdriver.crome.driver",  "C:\\DVGS\\SeleniumCrome\\chromedriver.exe");
	System.setProperty("webdriver.gecko.driver",  "C:\\DVGS\\SeleniumGecko\\geckodriver.exe");
	driver = new FirefoxDriver();
	driver.manage().window().maximize();
	driver.manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
	driver.get(baseUrl);
	String title = driver.getTitle();
	System.out.println(title);
	Assert.assertTrue(title.contains("OTM - Oracle Transportation Management"));
 
  }
  @DataProvider
  public String[][] getExcelData() throws InvalidFormatException, IOException {
	ReadExcel read = new ReadExcel();
	return read.getCellData("Testdata/XPOlogin.xls", "sheet1");
 }

  //@Test(dataProviderClass = com.abc.dataprovider.ExcelDataProvider.class,dataProvider="getData")
  //dataProviderClass = utilitiesXPO.ReadExcel.class, 
  @Test(dataProvider="getExcelData", priority = 0)
  public void testLoginOTM(String Userid, String Password) throws InterruptedException   {
	  WebDriverWait wait  = new WebDriverWait(driver,60);	
	  Thread.sleep(2000);
	  driver.findElement(By.xpath("//html/body/table/tbody/tr[2]/td/table/tbody/tr/td/form/table/tbody/tr[2]/td[2]/input")).clear();	  
	  driver.findElement(By.xpath("//html/body/table/tbody/tr[2]/td/table/tbody/tr/td/form/table/tbody/tr[2]/td[2]/input")).sendKeys(Userid);
	  driver.findElement(By.xpath("//html/body/table/tbody/tr[2]/td/table/tbody/tr/td/form/table/tbody/tr[3]/td[2]/input")).clear();	  
	  driver.findElement(By.xpath("//html/body/table/tbody/tr[2]/td/table/tbody/tr/td/form/table/tbody/tr[3]/td[2]/input")).sendKeys(Password);
	  driver.findElement(By.xpath("//html/body/table/tbody/tr[2]/td/table/tbody/tr/td/form/table/tbody/tr[4]/td[2]/input")).click();
	  Thread.sleep(2000);
	  String title = driver.getTitle();
	  System.out.println(title);
	  Assert.assertTrue(title.contains("XPO Logistics"));
  
  } 

 @AfterSuite
  public void ClosenOTM() {
	 System.out.println("Logout Successfull");
	 driver.close();
	 
  }

}
