import java.io.IOException;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import java.io.File;
import java.io.FileOutputStream;

public class Sheet
{

	public static void main(String[] args) throws IOException, InterruptedException
	{
		//Excel write code
		Workbook work = new XSSFWorkbook();
		org.apache.poi.ss.usermodel.Sheet sht = work.createSheet("Output_Sht");
		//product details
		String pageXpath[]= {"_35KyD6","_1vC4OE _3qQ9m1","hGSR34","VGWI6T _1iCvwn"};
		// call browser driver
		System.setProperty("webdriver.chrome.driver","D:\\New folder\\chromedriver.exe");
		//create object and launch driver
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.flipkart.com/");
		Thread.sleep(3000);
		//driver.manage().window().maximize();
		driver.findElement(By.xpath("//button[@class='_2AkmmA _29YdH8']")).click();
		driver.findElement(By.className("LM6RPg")).sendKeys("mobiles");
		driver.findElement(By.xpath("//button[@class='vh79eN']")).click();
		Thread.sleep(2000);		
		List<WebElement> mobile_count= driver.findElements(By.xpath("//div[@class='_3wU53n']"));
		String parGUID= driver.getWindowHandle();
		Actions builder = new Actions(driver);		
		int total=mobile_count.size();
		System.out.println("total:"+total);
		System.out.println(driver.getTitle());		
		for(int i=0; i<total;i++)
		{
			String frClick = "//div[@class='_3wU53n'][text()='";
			String mk = frClick + mobile_count.get(i).getText()+ "']";
			builder.moveToElement(driver.findElement(By.xpath(mk))).click().build().perform();
			
			for(String guid : driver.getWindowHandles()) 
			{
				if(!guid.equals(parGUID)) 
				{
					driver.switchTo().window(guid);
					Row newRow = sht.createRow(i);
					for(int j = 0;j<pageXpath.length;j++) 
					 {
						Cell cell = newRow.createCell(j);
						String forPass = "//*[@class = '"+pageXpath[j]+"']"; 
						try {
							cell.setCellValue(driver.findElement(By.xpath(forPass)).getText().toString());
						}
						catch(NoSuchElementException e) {
							cell.setCellValue("Nil");
						}
					 }
				}
			}
			driver.close();
			driver.switchTo().window(parGUID);
			
			//if(i==20) break;
		}
		FileOutputStream out = new FileOutputStream(new File("E:\\Nagamani.xlsx"));
		work.write(out);
		out.close();
		
		
	}
}
