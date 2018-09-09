package dynamicwebtable;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.Messaging.SyncScopeHelper;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;

public class Dynamictable 
{
	WebDriver driver;
	
	
	
	
  @Test
  public void Tabletest() throws AWTException, InterruptedException, IOException 
  {
		FileInputStream fi;
		FileOutputStream fo;
		XSSFWorkbook wb;
		XSSFSheet sh;


		String company = driver.findElement(By.xpath("//div[@id='leftcontainer']/child::table//th[1]")).getText();
		Robot rob = new Robot();
		rob.keyPress(KeyEvent.VK_ALT);
		rob.keyPress(KeyEvent.VK_SPACE);
		Thread.sleep(1000);
		rob.keyPress(KeyEvent.VK_N);
		rob.keyRelease(KeyEvent.VK_ALT);
		rob.keyRelease(KeyEvent.VK_BACK_SPACE);
		rob.keyRelease(KeyEvent.VK_N);
		Thread.sleep(1000);

		if (company.equalsIgnoreCase("Company")) {
			System.out.println("Company found");

			List<WebElement> l = driver
					.findElements(By.xpath("//div[@id='leftcontainer']/child::table//tbody/tr/td/a"));
			System.out.println("No. of companies= " + l.size());
			try {
				 	String heading = driver.findElement(By.xpath("//tr/th[text()='Company']")).getText();
					System.out.println("heading= "+heading);
					fi = new FileInputStream("E:\\automation\\data.xlsx");
				wb = new XSSFWorkbook(fi);
				sh = wb.getSheetAt(0);
				sh.createRow(0).createCell(0).setCellValue(heading);
				fo = new FileOutputStream("E:\\automation\\data.xlsx");
				wb.write(fo);
				
				for (int i = 1; i <= l.size(); i++) 
				{
				
					String s = driver.findElement(By.xpath("//tr/th[text()='Company']/ancestor::thead/following-sibling::tbody/tr["+i+"]/td[1]/a")).getText();
					System.out.println(i+" "+s);
					fi = new FileInputStream("E:\\automation\\data.xlsx");
					wb = new XSSFWorkbook(fi);
					sh = wb.getSheetAt(0);
					sh.createRow(i).createCell(0).setCellValue(s);
					fo = new FileOutputStream("E:\\automation\\data.xlsx");
					wb.write(fo);
					
				}

			} catch (FileNotFoundException e) {

				e.printStackTrace();
			}

		}

		else {
			System.out.println("Company not found");
		}
  }
  
  @BeforeTest
  public void beforeTest() 
  {
	  System.setProperty("webdriver.chrome.driver", "E:\\automation\\chromedriver\\chromedriver.exe");
	  driver= new ChromeDriver();
	  driver.get("https://money.rediff.com/gainers/bsc/daily/groupa");
  }

  @AfterTest
  public void afterTest() 
  {
	  System.out.println("-------execution complete----------");
	  driver.close();
  }

}
