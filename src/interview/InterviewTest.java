package interview;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class InterviewTest {

	WebDriver driver;

	@BeforeTest
	void beforeTest() {
		// Specify the path for geckodriver in below path.
		System.setProperty("webdriver.gecko.driver", "G:\\geckodriver.exe");
		driver = new FirefoxDriver();
	}

	@Test
	public void beforeCondition() throws IOException {

		// Specify the path for .xlsx excel file in below path.
		File file = new File("C:\\Users\\Daniel George\\Documents\\InterviewTest.xlsx");

		FileInputStream inputStream = new FileInputStream(file);

		Workbook wb = new XSSFWorkbook(inputStream);

		// Specify the sheet name below.
		Sheet currentSheet = wb.getSheet("Sheet1");

		int noOfRows = currentSheet.getLastRowNum();

		driver.get("http://www.google.com");

		boolean firstIter = true;

		for (int i = 0; i <= noOfRows; i++) {
			
			Row row = currentSheet.getRow(i);
			driver.findElement(By.xpath("//input[@id='lst-ib' and @title='Search']")).clear();
			driver.findElement(By.xpath("//input[@id='lst-ib' and @title='Search']")).sendKeys(row.getCell(0).toString());
			if (firstIter) {
				driver.findElement(By.xpath("//input[@name='btnK' and @value='Google Search']")).click();
				firstIter = false;
			} else {
				driver.findElement(By.xpath(
						"//input[@id='lst-ib' and @title='Search']/ancestor::div[@id='searchform']//button[@type='submit']"))
						.click();
			}
				
			String result = driver.findElement(By.xpath("//div[@id='resultStats']")).getText();
			String [] resultCount = result.split(" ");
			
			row.createCell(1).setCellValue("no of results :"+resultCount[1]);
			List<WebElement> elements = driver.findElements(
					By.xpath("//h2[contains(text(),'Search Results')]/following::div/div[@id='rso']//*[@href]"));
			int j = 2;
			for (WebElement element : elements) {
				String url = element.getAttribute("href");
				row.createCell(j).setCellValue(url);
				j++;
				if(j>6)
					break;
			}
		}
		inputStream.close();
		FileOutputStream fileOut = new FileOutputStream(file);
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
	
	@AfterTest
	void afterTestMethod() {
		driver.quit();
	}
}
