import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileOutputStream;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class Automation {

	public static void main(String[] args) throws Exception {
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\Naveenaa\\eclipse-workspace\\Selenium\\chromedriver\\chromedriver.exe");
		
		WebDriver driver = new ChromeDriver();
		
		Actions a = new Actions(driver);
		
		JavascriptExecutor js = (JavascriptExecutor) driver;
		
		TakesScreenshot tk = (TakesScreenshot) driver;
		
		driver.get("https://www.shoespie.com/");
		
		driver.manage().window().maximize();
		
		//driver.switchTo().alert().dismiss();
		
		Thread.sleep(2500);
		
		//SCREENSHOT-1
		
		File temp0 = tk.getScreenshotAs(OutputType.FILE);
		
		File dest0 = new File ("C:\\Users\\Naveenaa\\eclipse-workspace\\MiniProject\\Screenshot\\screen1.jpeg");
		
		FileUtils.copyFile(temp0, dest0);
		
		//
		
		driver.findElement(By.xpath("//button[@id='onesignal-popover-cancel-button']")).click();
					
		WebElement sandals = driver.findElement(By.xpath("//li[@class='cat-item Sandals']"));
		
		WebElement flats = driver.findElement(By.xpath("//a[text()='Flat Sandals']"));
			
		a.moveToElement(sandals).perform();
		
		js.executeScript("arguments[0].click()", flats);
		
		//SCREENSHOT-1
		
		File temp = tk.getScreenshotAs(OutputType.FILE);
		
		File dest = new File ("C:\\Users\\Naveenaa\\eclipse-workspace\\MiniProject\\Screenshot\\screen2.jpeg");
		
		FileUtils.copyFile(temp, dest);
		
		//
		
		Robot r = new Robot();
		
		r.keyPress(KeyEvent.VK_PAGE_DOWN);
		r.keyRelease(KeyEvent.VK_PAGE_DOWN);
		
		WebElement product = driver.findElement(By.xpath("//a[@class='tidely gary']"));
		
		js.executeScript("arguments[0].click()",product);
		
	//SCREENSHOT-2
		
		File temp1 = tk.getScreenshotAs(OutputType.FILE);
		
		File dest1 = new File ("C:\\Users\\Naveenaa\\eclipse-workspace\\MiniProject\\Screenshot\\screen2.jpeg");
		
		FileUtils.copyFile(temp1, dest1);
		
		//
		
		WebElement title = driver.findElement(By.xpath("//h1[text()='Shoespie Trendy Flat Flip Flop Rhinestone Glitter Slippers']"));
		
		String toExcel = title.getText();
		
		WebElement price = driver.findElement(By.xpath("//span[@class='multi']"));
		
		String toExcel1 = price.getText() ;
		
		driver.findElement (By.xpath("//li[@data-valueid='299']")).click();//color
		
		driver.findElement(By.xpath("//div[@class='required-txt']")).click();//size
		
		driver.findElement(By.xpath("//span[@data-pair='12=338']")).click();
		
		WebElement btnToCart = driver.findElement(By.id("addBtn"));
		
		Thread.sleep(750);
		
		js.executeScript("arguments[0].click()", btnToCart);
		
		Thread.sleep(750);
		
		js.executeScript("arguments[0].click()", btnToCart);
				
		File loc = new File ("C:\\Users\\Naveenaa\\eclipse-workspace\\MiniProject\\Excel\\naveenaa.xlsx");
		
		Workbook w = new XSSFWorkbook();
		
		Sheet sh = w.createSheet("Test data");
		
		Row r1 = sh.createRow(0);
		
		Cell c = r1.createCell(0);
		
		c.setCellValue("Product-name");
		
		Cell c1 = r1.createCell(1);
		
		c1.setCellValue("Price");
		
		Row r2 = sh.createRow(1);
		
		Cell c3 = r2.createCell(0);
		
		c3.setCellValue(toExcel);
		
		Cell c4 = r2.createCell(1);
		
		c4.setCellValue(toExcel1);
		
		FileOutputStream o = new FileOutputStream(loc);
		
		w.write(o);
		
		System.out.println("Excel updated");
		
		//SCREENSHOT-3
		
				File temp2 = tk.getScreenshotAs(OutputType.FILE);
				
				File dest2 = new File ("C:\\Users\\Naveenaa\\eclipse-workspace\\MiniProject\\Screenshot\\screen3.jpeg");
				
				FileUtils.copyFile(temp2, dest2);
				//		
		driver.quit();
		
	}
	
	
}
