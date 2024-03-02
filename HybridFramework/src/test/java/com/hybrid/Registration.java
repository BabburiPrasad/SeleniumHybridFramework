package com.hybrid;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.security.SecureRandom;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Registration 
{
	static WebDriver driver;
	static XSSFWorkbook wbook;
	static XSSFSheet sheet;
	static String baseURL = "https://naveenautomationlabs.com/opencart/index.php?route=account/register";

	public static String getCellValues(int row, int col)
	{
		XSSFRow r = sheet.getRow(row);
		XSSFCell c = r.getCell(col);
		
		String a = c.getStringCellValue();
		
		return a;
	}
	@SuppressWarnings("deprecation")
	public static void main(String[] args) throws InterruptedException 
	{
		int length = 10; // Change the length as per your requirement
        String randomString = generateRandomString(length);
        String domain = "example.com"; // Change the domain name as per your requirement
        String email = randomString + "@" + domain;
        System.out.println("Random Email: " + email);
		
		
		File fi = new File("C:\\Users\\HP\\Documents\\RegTestData.xlsx");
		try 
		{
			FileInputStream fis = new FileInputStream(fi);
			wbook = new XSSFWorkbook(fis);
			
			sheet = wbook.getSheet("Data");
		} 
		catch (IOException e) 
		{
			System.out.println("File Not Found, Please check the file path: "+e);
		}
		
		int rows = sheet.getPhysicalNumberOfRows();
		
		for(int i = 0; i<rows; i++)
		{
			String action  = getCellValues(i, 2);
			
			System.out.println(action);
			switch (action)
			{
				case "Open":
					driver = new ChromeDriver();
					driver.manage().window().maximize();
					driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
					
				break;
				
				case "Navigate":
					driver.get(baseURL);
				break;
				
				case "TypeF":
					// Fetches FirstName text-box xpath and first-name value from xlsx file
					driver.findElement(By.xpath(getCellValues(2, 3))).sendKeys(getCellValues(2, 4));
					Thread.sleep(3000);
				break;
					
				case "TypeL":
				
					// Fetches LastName text-box xpath and Last-name value from xlsx file
					driver.findElement(By.xpath(getCellValues(3, 3))).sendKeys(getCellValues(3, 4));
					Thread.sleep(3000);	
				break;
				
				// Fetches Email text-box xpath from xlsx file and takes random email from generateRandomString()
				case "TypeE":
					driver.findElement(By.xpath(getCellValues(4, 3))).sendKeys(email);
					Thread.sleep(3000);
					
				break;
				
				case "TypePh":
					
					// Fetches Phone-Number text-box xpath and Phone-Number value from xlsx file
					driver.findElement(By.xpath(getCellValues(5, 3))).sendKeys("9874563210");
					Thread.sleep(3000);
					
				break;
				case "TypeP":
					
					// Fetches Password text-box xpath and Password value from xlsx file
					driver.findElement(By.xpath(getCellValues(6, 3))).sendKeys(getCellValues(6, 4));
					Thread.sleep(3000);
					
				break;
				
				case "TypeCP":
					
					// Fetches CNF-Password text-box xpath and CNF-Password value from xlsx file
					driver.findElement(By.xpath(getCellValues(7, 3))).sendKeys(getCellValues(7, 4));
					Thread.sleep(3000);
					
				break;
					
				case "ClickR":
					// Clicks on Subscribe link
					driver.findElement(By.xpath(getCellValues(8, 3))).click();
					Thread.sleep(3000);
				break;
				
				case "ClickC":
					// Clicks on Privacy Policy check-box
					driver.findElement(By.xpath(getCellValues(9, 3))).click();
					Thread.sleep(3000);
				break;
				
				case "ClickCB":
					// Clicks on Privacy Policy check-box
					driver.findElement(By.xpath(getCellValues(10, 3))).click();
					Thread.sleep(3000);
				break;
				
				default: System.out.println("No Action Found");
				
			
			}
			
			
		}

	}
	
	private static final String CHARACTERS = "abcdefghijklmnopqrstuvwxyz0123456789";
    private static final SecureRandom random = new SecureRandom();
	
	public static String generateRandomString(int length) {
        StringBuilder stringBuilder = new StringBuilder(length);
        for (int i = 0; i < length; i++) {
            stringBuilder.append(CHARACTERS.charAt(random.nextInt(CHARACTERS.length())));
        }
        return stringBuilder.toString();
    }

}
