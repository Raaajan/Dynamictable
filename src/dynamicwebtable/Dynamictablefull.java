package dynamicwebtable;

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
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Dynamictablefull {


	public static void main(String[] args) throws IOException, AWTException, InterruptedException {
		

		FileInputStream fi;
		FileOutputStream fo;
		XSSFWorkbook wb;
		XSSFSheet sh;
		System.setProperty("webdriver.chrome.driver", "E:\\automation\\chromedriver\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("https://money.rediff.com/gainers/bsc/daily/groupa");

		
		Robot rob = new Robot();
		rob.keyPress(KeyEvent.VK_ALT);
		rob.keyPress(KeyEvent.VK_SPACE);
		Thread.sleep(1000);
		rob.keyPress(KeyEvent.VK_N);
		rob.keyRelease(KeyEvent.VK_ALT);
		rob.keyRelease(KeyEvent.VK_BACK_SPACE);
		rob.keyRelease(KeyEvent.VK_N);
		Thread.sleep(1000);

		List<WebElement> columncount = driver.findElements(By.xpath("//div[@id='leftcontainer']/child::table//thead/tr/th"));
		System.out.println("total columns = "+columncount.size());
		
		fi = new FileInputStream("E:\\automation\\data1.xlsx");
		wb= new XSSFWorkbook(fi);
		sh= wb.getSheetAt(0);
		sh.createRow(0);
		for(int i=1;i<=columncount.size();i++)
		{
			String columnname=driver.findElement(By.xpath("//div[@id='leftcontainer']/child::table//thead/tr/th["+i+"]")).getText();
			System.out.println("columnname "+columnname );
			sh.getRow(0).createCell(i).setCellValue(columnname);
			fo= new FileOutputStream("E:\\automation\\data1.xlsx");
			wb.write(fo);
		}
		
		
		List <WebElement> rowcount = driver.findElements(By.xpath("//div[@id='leftcontainer']/table//tbody/tr"));
		System.out.println("row count= "+rowcount.size());
		
		for(int i=1;i<=rowcount.size();i++)
		{	sh.createRow(i);
			for(int j=1;j<=columncount.size();j++)
			{
				String data = driver.findElement(By.xpath("//div[@id='leftcontainer']/table//tbody/tr["+i+"]/td["+j+"]")).getText();
				System.out.println(data);
				sh.getRow(i).createCell(j).setCellValue(data);
				fo= new FileOutputStream("E:\\automation\\data1.xlsx");
				wb.write(fo);
			}
		}
		
		System.out.println();
		System.out.println("------execution complete----------");

		driver.close();
	}

}
