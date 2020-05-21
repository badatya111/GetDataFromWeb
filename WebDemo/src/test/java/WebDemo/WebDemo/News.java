package WebDemo.WebDemo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class News extends BrowserLaunch{
	

	/*public News(String newsData, String strUrl) {
		super(newsData, strUrl);
		// TODO Auto-generated constructor stub
	}
	//public News(String browser) {
		//super(newsData, browser);
		// TODO Auto-generated constructor stub
	//}
	public static News news;
	//public static News get() 
	//{ 
		//if (news == null) 
			//news = new News(newsData); 
		//return news; 
	//} 
	public void newsData(int totalNoOfPage,String sheetName,String excelFileName) throws IOException, EncryptedDocumentException, InvalidFormatException 
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
				//File fi=new File("login.xlsx");
				//fi.createNewFile();
				//Economy.xlsx
				FileInputStream inputStream = new FileInputStream(new File(excelFileName));
				Workbook workbook = WorkbookFactory.create(inputStream);
				//XSSFWorkbook workbook = new XSSFWorkbook(); 
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
				js.executeScript("javascript:window.scrollBy(130,130)");
				driver.findElement(By.xpath("(//*[@class='last']/span)[1]")).click();
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
		}
	} 
	
*/
}
