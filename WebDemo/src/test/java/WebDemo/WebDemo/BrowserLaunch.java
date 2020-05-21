package WebDemo.WebDemo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class BrowserLaunch extends Thread{

	public static WebDriver driver;
	public  String newsData;
	public  String strUrl;


	/*public BrowserLaunch(String newsData,String strUrl){
		this.strUrl=strUrl;
		this.newsData=newsData;
	}
	public void run(){
		String strCrrentThread=Thread.currentThread().getName();
		System.out.println("Current Thread Name :"+strCrrentThread);
		try {
			launchApp(newsData,strUrl);
			driver.close();
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}*/
	public static  void newsData(int totalNoOfPage,String sheetName,String excelFileName) throws IOException, EncryptedDocumentException, InvalidFormatException 
	{ 
		try {
			Thread.sleep(8000);
		} catch (InterruptedException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		// This data needs to be written (Object[]) 
		for(int j=1;j<=totalNoOfPage;j++){
			String strUrl=driver.getCurrentUrl();
			try {
				Thread.sleep(500);
				List<String[]> data = new LinkedList< String[]>(); 
				//data.add(new String[]{"Date", "HeadLine", "Subject","strUrl"}); 
				WebDriverWait wait = new WebDriverWait (driver, 60);
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='cagetory']/li/span"))); 
				List<WebElement> links = driver.findElements(By.xpath("//*[@id='cagetory']/li/span"));
				int b=links.size();
				System.out.println("Loop count :"+j+"  : Size : "+b);
				String strDate,strHeader,strSubject;
				String strDateList = null,headLine = null,subject = null;
				int i=1;
				for(;i<b;i++){
					strDate="(//*[@id='cagetory']/li/span)["+i+"]";
					strHeader="(//*[@id='cagetory']/li/h2)["+i+"]";
					strSubject="(//*[@id='cagetory']/li/p)["+i+"]";
					strDateList=driver.findElement(By.xpath(strDate)).getText();
					headLine=driver.findElement(By.xpath(strHeader)).getText();
					subject=driver.findElement(By.xpath(strSubject)).getText();
					data.add(new String[]{ strDateList, headLine,subject,strUrl}); 
				}

				FileInputStream inputStream = new FileInputStream(new File(excelFileName));
				Workbook workbook = WorkbookFactory.create(inputStream);

				// Create a blank sheet 
				Sheet sheet = (workbook.getSheet(sheetName)!=null)?workbook.getSheet(sheetName):workbook.createSheet(sheetName); 
				int rownum = sheet.getLastRowNum() ; 
				for (String objArr[] : data) { 
					Row row = sheet.createRow(rownum++); 
					int cellnum = 0; 
					for (String obj : objArr) { 
						Cell cell = row.createCell(cellnum++); 
						cell.setCellValue(obj); 
					} 
				} 
				try { 
					FileOutputStream out = new FileOutputStream(new File(excelFileName)); 
					workbook.write(out); 
					workbook.close();
					out.close();
					System.out.println("written successfully on disk."); 
				} 
				catch (Exception e) { 
					e.printStackTrace(); 
				} 
				JavascriptExecutor js = (JavascriptExecutor) driver;
				js.executeScript("javascript:window.scrollBy(132,132)");
				//driver.findElement(By.xpath("(//*[@class='last']/span)[1]")).click();
				js.executeScript("arguments[0].click();", driver.findElement(By.xpath("(//*[@class='last']/span)[1]")));
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	} 
	public static void launchApp( ) throws EncryptedDocumentException, InvalidFormatException, IOException{
		//String browser="browser";
		String strUrl="https://www.moneycontrol.com/news/business/stocks/page-1645/";
		String currentDirectory = System.getProperty("user.dir");
		String chromeDiver;
		System.out.println("Chrome Driver Launch");
		chromeDiver=currentDirectory+"\\Drivers\\chromedriver.exe";
		System.out.println(chromeDiver);
		System.setProperty("webdriver.chrome.driver", chromeDiver);
		ChromeOptions op = new ChromeOptions(); 
		Map<String, Object> prefs = new HashMap<>();
		op.addArguments("headless");
		op.addArguments("window-size=1200x600");
		prefs.put("profile.default_content_setting_values.notifications", 1); 
		op.setExperimentalOption("prefs", prefs); 
		driver=new ChromeDriver(op);
		driver.get(strUrl);
		driver.manage().window().maximize();
		newsData(5922, "stock", "Stock.xlsx");
	}

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		/*Thread t1=new BrowserLaunch("business","https://www.moneycontrol.com/news/business/");
		Thread t2=new BrowserLaunch("market","https://www.moneycontrol.com/news/business/markets/");
		Thread t3=new BrowserLaunch("stocks","https://www.moneycontrol.com/news/business/stocks/");
		Thread t4=new BrowserLaunch("pf","https://www.moneycontrol.com/news/business/personal-finance/");
		Thread t5=new BrowserLaunch("mfund","https://www.moneycontrol.com/news/business/mutual-funds/");
		t1.start();
		t2.start();
		t3.start();
		t4.start();
		t5.start();*/
		launchApp();
	}
}
