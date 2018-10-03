package testScript;

import java.util.Calendar;
import java.util.Date;
import java.util.concurrent.TimeUnit;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;
//import testScript.TC001_LoginOTM;

import com.sun.rowset.internal.Row;

import utilitiesXPO.ReadExcel;

public class FastOrderEntry {
 // private WebDriver driver;
//	TC001_LoginOTM tc = new TC001_LoginOTM();
//	  WebDriver wDriver = (WebDriver)tc.driver;
	@DataProvider
	public String[][] getExcelData() throws InvalidFormatException, IOException {
		ReadExcel read = new ReadExcel();
		return read.getCellData("Testdata/FOE.xlsx", "sheet1");
	 }
	@Test(dataProvider="getExcelData", priority=1)
	 
	  public void clientQuickLinks(String OrigionLocation, String Dest, String Item, String Count, String Weight, String Volume, String Client_Order, String Busunit, String GeneralLedger) throws InterruptedException, IOException   {
		 //Integer i = Integer.getInteger(Item);
		
	 	 /* WebDriverWait wait  = new WebDriverWait(TC001_LoginOTM.getDriver(),80);	
	 	  wait.until(ExpectedConditions.elementToBeClickable(element).)*/
		TC001_LoginOTM.getDriver().manage().timeouts().implicitlyWait(30,TimeUnit.SECONDS);
		  
		  TC001_LoginOTM.getDriver().switchTo().defaultContent();
		  System.out.println("before sidebar");
		  TC001_LoginOTM.getDriver().switchTo().frame("sidebar");
		  System.out.println("After sidebar");
		  //click on client quick links 
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//a[contains(text(),'Client Quick Links')]")).click(); 	  
		  
		  //click on FOE screen
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//a[text()='Fast Order Entry']")).click();
		  		  
		  //Internal Login screen
		  TC001_LoginOTM.getDriver().switchTo().defaultContent();
		  TC001_LoginOTM.getDriver().switchTo().frame("mainBody");
				  
		  //Internal username
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//*[@id='orderentryui-1928791928']/div/div[2]/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/table/tbody/tr[1]/td[3]/input")).clear();
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//*[@id='orderentryui-1928791928']/div/div[2]/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/table/tbody/tr[1]/td[3]/input")).sendKeys("MENLO.CLIENTA370");
		  
		  //Internal password
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//*[@id=\"orderentryui-1928791928\"]/div/div[2]/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/table/tbody/tr[2]/td[3]/input")).clear();
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//*[@id=\"orderentryui-1928791928\"]/div/div[2]/div/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/table/tbody/tr[2]/td[3]/input")).sendKeys("Year2012??");
		  
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[1]/div/div/div/div[2]/div/div/div/div[3]/div/div/span/span")).click();
         
		  
		  //fast order entry screen
		  TC001_LoginOTM.getDriver().switchTo().defaultContent();
		  TC001_LoginOTM.getDriver().switchTo().frame("mainBody");
		  WebElement Origion = TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[2]/div/div/div/div[2]/div/input"));
		  Origion.sendKeys(OrigionLocation);
		  Thread.sleep(5000);
		  //Origion.sendKeys(Keys.TAB);
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[3]/div/div/div/div[2]/div/input")).sendKeys(Dest);
		  Thread.sleep(5000);
		  
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[7]/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[1]/div/div/div/div[1]/div/input")).sendKeys(Item);
		  Thread.sleep(5000);
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div[3]/div/div/div/div[3]/div/div/div[2]/div/div/div[1]/div/div/div[2]/div[1]/table/tbody/tr[1]/td[1]/div")).click();
		  Thread.sleep(8000);
		  		  
		  //to enter count
		                
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[7]/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[1]/div/div/div/div[4]/div/input")).sendKeys(Count);
		  Thread.sleep(2000);
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[7]/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[1]/div/div/div/div[5]/div/input")).sendKeys(Weight);
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[7]/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[1]/div/div/div/div[7]/div/input")).sendKeys(Volume);
		  Thread.sleep(9000);
		  //changing date format
		  //Create object of SimpleDateFormat class and decide the format
		  DateFormat dateFormat = new SimpleDateFormat("M/d/yy");
		  //get current date time with Date()
		  Date date = new Date();
		  // Now format the date
		  String date1= dateFormat.format(date);
		  // to add 4 days to codays date
		  Calendar cal = Calendar.getInstance();
		  cal.setTime(date);
		  cal.add(Calendar.DATE, 4); //minus number would decrement the days
		  Date  date2 = cal.getTime();
		  String date3 = dateFormat.format(date2);		  
		  		  		  
		  //entering todays date fields
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[9]/div/div/div/div[2]/div/div/div/div[2]/div/div/input")).sendKeys(date1);
		  // entering todays date +4 days
		  Thread.sleep(2000);
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[10]/div/div/div/div[3]/div/div/div/div[2]/div/div/input")).sendKeys(date3);
		  Thread.sleep(2000);
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[12]/div/div/div/div/div[2]/div/input")).sendKeys(Client_Order);
		  Thread.sleep(2000);
		  //selecting from combo box
		  System.out.println("before selection element");
		  WebElement BusUnitElement = TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[12]/div/div/div/div/div[6]/div/div/input"));
		  BusUnitElement.sendKeys(Busunit);
	  
		  //Select BusUnit = new Select(BusUnitElement);
		  //BusUnit.selectByVisibleText("TRS");
		  System.out.println("After selection element");
		  Thread.sleep(5000);
		  System.out.println("before premium check box element");
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//*[@id=\"gwt-uid-3\"]")).click();
		  
		  //select premium reason from combo box
		  Thread.sleep(2000);
		  WebElement PreReasonElement = TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[12]/div/div/div/div/div[22]/div/div/select"));
		  Select PreReason = new Select(PreReasonElement);
		  PreReason.selectByIndex(2);
		  
		  Thread.sleep(9000);
		  
		  //General Ledger from combo box
		  WebElement GenLedElement = TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[15]/div/div/div/div[2]/div/div/input"));
		  //Select GenLed = new Select(GenLedElement);
		  //GenLed.selectByVisibleText("143-16400-000");
		  GenLedElement.sendKeys(GeneralLedger);
		  Thread.sleep(5000);
		  GenLedElement.sendKeys(Keys.ARROW_DOWN.ENTER);
		   
		  Thread.sleep(9000);
		  
		  //click on create button
		  TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[25]/div/div/div/div[1]/div/div/div/div[1]/div/div/div/div[1]/div/div/span/span")).click();
		  Thread.sleep(9000);
		  
		  
		  
		  //String message = "//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[25]/div/div/div/div[1]/div/div/div/div[2]/div/div/div/div[1]/div/div";
		  System.out.println("Before popup window");
		  new WebDriverWait(TC001_LoginOTM.getDriver(),30).until(ExpectedConditions.elementToBeClickable(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[25]/div/div/div/div[1]/div/div/div/div[2]/div/div")));
		  
		  System.out.println("After popup window");
		  Thread.sleep(9000);
		  		  
		  
		  String element_text = TC001_LoginOTM.getDriver().findElement(By.xpath("//html/body/div/div/div[2]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/div/div[25]/div/div/div/div[1]/div/div/div/div[2]/div/div")).getText();
		  //splitting on the basis of space and taking index 1 to take order no 
		  String element_text_actual = element_text.split(" ")[1];
		  System.out.println(element_text_actual);	
		 
		
		 try {
			 //writing order no to Excel file
			  // HSSFWorkbook wb = new HSSFWorkbook();   
			  //CreationHelper createHelper = wb.getCreationHelper();
			  //Sheet sheet = wb.createSheet("Sheet1");
			  
			    //HSSFRow row = (HSSFRow) sheet.createRow(1);
		        //HSSFCell cell;
		        //Creating rows and filling them with data 
		       // cell = row.createCell(1);
		       // cell.setCellValue(createHelper.createRichTextString(element_text_actual));
			 // FileOutputStream fileOut;
             // fileOut = new FileOutputStream("Testdata/FOE_output.xls");
			 // wb.write(fileOut);
			 // fileOut.close(); 
			//create an object of Workbook and pass the FileInputStream object into it to create a pipeline between the sheet and eclipse.
				FileInputStream fis = new FileInputStream("Testdata/FOE_output.xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				//call the getSheet() method of Workbook and pass the Sheet Name here. 
				//In this case I have given the sheet name as “TestData” 
		                //or if you use the method getSheetAt(), you can pass sheet number starting from 0. Index starts with 0.
				XSSFSheet sheet = workbook.getSheet("sheet1");
				//XSSFSheet sheet = workbook.getSheetAt(0);
				//Now create a row number and a cell where we want to enter a value. 
				//Here im about to write my test data in the cell B2. It reads Column B as 1 and Row 2 as 1. Column and Row values start from 0.
				//The below line of code will search for row number 2 and column number 2 (i.e., B) and will create a space. 
		                //The createCell() method is present inside Row class.
		                XSSFRow row = sheet.createRow(1);
				Cell cell = row.createCell(1);
				//Now we need to find out the type of the value we want to enter. 
		                //If it is a string, we need to set the cell type as string 
		                //if it is numeric, we need to set the cell type as number
				cell.setCellType(cell.CELL_TYPE_STRING);
				cell.setCellValue(element_text_actual);
				FileOutputStream fos = new FileOutputStream("Testdata/FOE_output.xlsx");
				workbook.write(fos);
				fos.close();
				System.out.println("END OF WRITING DATA IN EXCEL");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
		}  
		
		 Assert.assertTrue(element_text.contains("Order"));		  
		  
	  }
}
