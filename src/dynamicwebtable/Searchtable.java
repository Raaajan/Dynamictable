package dynamicwebtable;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeTest;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;

import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterTest;

public class Searchtable {
	WebDriver driver;
	FileInputStream fi;
	FileOutputStream fo;
	XSSFWorkbook wb;
	XSSFSheet sh;
	
	@Test
	public void Tabletest() throws AWTException, InterruptedException, IOException {

		Robot rob = new Robot();
		rob.keyPress(KeyEvent.VK_ALT);
		rob.keyPress(KeyEvent.VK_SPACE);
		Thread.sleep(1000);
		rob.keyPress(KeyEvent.VK_N);
		rob.keyRelease(KeyEvent.VK_ALT);
		rob.keyRelease(KeyEvent.VK_BACK_SPACE);
		rob.keyRelease(KeyEvent.VK_N);
		Thread.sleep(1000);

		List<WebElement> columns = driver.findElements(By.xpath("//div[@id='leftcontainer']/child::table//thead/tr/th"));
		int cc=columns.size();
		System.out.println("total columns ="+cc);
		
		String input="Company";
		
		for(int i=1;i<=cc;i++)
		{
			String header=driver.findElement(By.xpath("//div[@id='leftcontainer']/child::table//thead/tr/th["+i+"]")).getText();
			System.out.println("header= "+header);
			
			if(header.equalsIgnoreCase(input))
			{
				List<WebElement> rows = driver.findElements(By.xpath("//div[@id='leftcontainer']/child::table//tbody/tr"));
				int rc = rows.size();
				System.out.println("total rows= "+rc);
				
				for(int j=1;j<=rc;j++)
				{
					String rowval=driver.findElement(By.xpath("//div[@id='leftcontainer']/child::table//tbody/tr["+j+"]/td["+i+"]")).getText();
					System.out.println(j+" "+rowval);
				
					fi = new FileInputStream("E:\\automation\\searchtabledata.xlsx");
					wb= new XSSFWorkbook(fi);
					sh= wb.getSheetAt(0);
				//	sh.createRow(0).createCell(0).setCellValue(header);
					sh.createRow(j).createCell(0).setCellValue(rowval);
					fo= new FileOutputStream("E:\\automation\\searchtabledata.xlsx");
					wb.write(fo);
				}
				break;
			}
			
		}
	}

	@BeforeTest
	public void beforeTest() {
		System.setProperty("webdriver.chrome.driver", "E:\\automation\\chromedriver\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.get("https://money.rediff.com/gainers/bsc/daily/groupa");
	}

	@AfterTest
	public void afterTest() {
		System.out.println("-------execution complete----------");
		driver.close();
	}

}
