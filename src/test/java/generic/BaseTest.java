package generic;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.ITestResult;
import org.testng.Reporter;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;

public class BaseTest {
	public WebDriver driver;
	public WebDriverWait wait;
	@Parameters({"grid","gridURL","browser","appURL","ITO","ETO"})
	@BeforeMethod
	public void preCondition(
			@Optional("no") String grid,
			@Optional("http://192.168.0.120:4444") String gridURL,
			@Optional("chrome") String browser, 
			@Optional("https://opensource-demo.orangehrmlive.com") String appURL,
			@Optional("10") String ITO,
			@Optional("10") String ETO
			) throws MalformedURLException
	{
		if(grid.equalsIgnoreCase("yes"))
		{
			Reporter.log("URL=" + gridURL , true);
			URL url=new URL(gridURL);
			if(browser.equalsIgnoreCase("chrome"))
			{
			
			ChromeOptions options=new ChromeOptions();
			driver=new RemoteWebDriver(url,options);
			}
			else
			{
				
				EdgeOptions options=new EdgeOptions();
				driver=new RemoteWebDriver(url,options);
			}
			
		}
		else
		{
			if(browser.equalsIgnoreCase("chrome"))
			{
				Reporter.log("browser=chrome" , true);
			driver=new ChromeDriver();
			}
			else
			{   Reporter.log("browser=Edge" , true);
				driver=new EdgeDriver();
			}
			
		}
		Reporter.log("Preconditionmethod",true);
		
		
		Reporter.log("Set ITO"+ ITO,true);
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(Integer.valueOf(ITO)));
		//driver=new ChromeDriver();
		Reporter.log("EnterURL",true);
		driver.get(appURL);
		Reporter.log("Maximise the browser",true);
		driver.manage().window().maximize();
		Reporter.log("Set ETO"+ ETO,true);
		wait=new WebDriverWait(driver, Duration.ofSeconds(Integer.valueOf(ETO)));
		
	}
	@AfterMethod
	public void postCondition(ITestResult result) throws Exception
	{
		String name = result.getName();
		int status = result.getStatus();
	
		if(status==2)
		{
			SimpleDateFormat dateFormat = new SimpleDateFormat("_yyyy_MM_dd_HH_mm_ss");
	        String timestamp = dateFormat.format(new Date());
			TakesScreenshot t=(TakesScreenshot) driver;
			File scrImage = t.getScreenshotAs(OutputType.FILE);
			File dstImage= new File("./images/"+name+timestamp+".png");
			FileUtils.copyFile(scrImage, dstImage);
					
		}
		
		
		Reporter.log("Close the browser",true);
		driver.quit();
	}

}
