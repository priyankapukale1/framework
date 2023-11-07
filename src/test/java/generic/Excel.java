package generic;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excel {

	public static String getData(String path,String sheet,int r,int c)
	{
		String v="";
		try 
		{
			Workbook wb = WorkbookFactory.create(new FileInputStream(path));
			v=wb.getSheet(sheet).getRow(r).getCell(c).toString();
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		
		return v;
	}
	
	public static int getRowCount(String path,String sheet)
	{
		int rowCount=0;
		
		try
		{
			Workbook wb = WorkbookFactory.create(new FileInputStream(path));
			rowCount=wb.getSheet(sheet).getLastRowNum();
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		
		return rowCount;
	}
	public static int getCellCount(String path,String sheet,int c1)
	{
		int cellCount=0;
		
		try
		{
			Workbook wb = WorkbookFactory.create(new FileInputStream(path));
			cellCount=wb.getSheet(sheet).getRow(c1).getLastCellNum();
			
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		
		return cellCount;
	}
	public static String setData(String path,String sheet,int r,int c,String s)
	{
		String v1="";
		try 
		{
			Workbook wb = WorkbookFactory.create(new FileInputStream(path));
			wb.getSheet(sheet).getRow(r).getCell(c).setCellValue(s);
			wb.write(new FileOutputStream(path));
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		
		return v1;
	}
	
	
	//add a method to count columns
	
	//method to write the data
	
}
