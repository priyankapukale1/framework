package script;

import org.testng.Assert;
import org.testng.Reporter;
import org.testng.annotations.Test;

import generic.BaseTest;
import generic.Excel;

public class DemoTest extends BaseTest
{

	@Test
	public void testDemo()
	{
		int r = Excel.getRowCount("./data/input.xlsx", "Sheet1");
		Reporter.log("Rowcount:"+r,true);
		int c = Excel.getCellCount("./data/input.xlsx","Sheet1",2);
		Reporter.log("Columncount:"+c,true);
		
		String d = Excel.getData("./data/input.xlsx", "Sheet1",2,2);
		
		Reporter.log("Data:"+d,true);
		String s=Excel.setData("./data/input.xlsx", "Sheet1",2,2,"clone");
		Reporter.log("Data:"+s,true);
		
		
		Reporter.log("test demo...",true);
		//Assert.fail();
	}
}