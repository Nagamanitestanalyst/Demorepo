package exelwrite;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
//_35KyD68888

public class Exel 
{
	 
	 public static void main (String [] args) throws InterruptedException, IOException
	 {
		 //Create excel sheet
		 Workbook work = new XSSFWorkbook();
			Sheet sht = work.createSheet("mobiles");
			
			String pageXpath[]= {"_35KyD6","_1vC4OE _3qQ9m1","hGSR34","_2-n-Lg col"};  
			System.setProperty("webdriver.chrome.driver","D:\\New folder\\chromedriver.exe");
			//initialize web driver
			WebDriver driver = new ChromeDriver();
			driver.get("https://www.flipkart.com/");//navigate URL
			driver.findElement(By.xpath("//button[@class='_2AkmmA _29YdH8']")).click();//To quit popup
			driver.findElement(By.className("LM6RPg")).sendKeys("mobiles");//find mobiles
			driver.findElement(By.xpath("//button[@class='vh79eN']")).click();
			String parGUID= driver.getWindowHandle();
			Thread.sleep(1000);	
			int rowNum = 0;
			int i;
			// find next page	
			List<WebElement> mobile_count;
			for(int k=1;k<=2;k++) {
				System.out.println(k);
				Thread.sleep(3000);
			mobile_count= driver.findElements(By.xpath("//div[@class='_3wU53n']"));
			//String parGUID= driver.getWindowHandle();
			Actions builder = new Actions(driver);		
			int total=mobile_count.size();
			System.out.println("total:"+total);
			System.out.println(driver.getTitle());
			Row newRow = sht.createRow(0);
			newRow.createCell(0).setCellValue("Name");
			newRow.createCell(1).setCellValue("Price");
			newRow.createCell(2).setCellValue("Rating");
			newRow.createCell(3).setCellValue("Offers");
			//System.out.println("enter value:");
			for( i=0; i<total;i++)
			{
				String frClick = "//div[@class='_3wU53n'][text()='";
				String mk = frClick + mobile_count.get(i).getText()+ "']";
				
				builder.moveToElement(driver.findElement(By.xpath(mk))).click().build().perform();
				
				for(String guid : driver.getWindowHandles()) 
				{
					if(!guid.equals(parGUID)) 
					{
						driver.switchTo().window(guid);
						Row newRow1 = sht.createRow(rowNum++);
						for(int j = 0;j<pageXpath.length;j++) 
						 {
							Cell cell = newRow1.createCell(j);
							String forPass = "//*[@class = '"+pageXpath[j]+"']"; 
							
							try {
								cell.setCellValue(driver.findElement(By.xpath(forPass)).getText().toString());	}
							catch(NoSuchElementException e) 
							{
							cell.setCellValue("Nil");
							} }
					
				driver.close();
				driver.switchTo().window(parGUID);
				//if(i==5) break;			
			}				} |}
		
if(i-1==total-1)
	((JavascriptExecutor) driver).executeScript("arguments[0].click();", driver.findElement(By.xpath(""
			+ "//a[@class='_2Xp0TH'][text()='"+String.valueOf(k+1)+"']")));
mobile_count.clear();

			}
			FileOutputStream out = new FileOutputStream(new File("E:\\Flipkartmobiles.xlsx"));
			work.write(out);
			out.close();
			
						
		}
	}