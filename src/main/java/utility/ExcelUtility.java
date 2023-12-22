package utility;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {
	
	public static FileInputStream fis;
	public static FileOutputStream fos;
	public static XSSFWorkbook wb;
	public static XSSFSheet sheet;
	public static XSSFRow row;
	public static XSSFCell cell;
	final int rowCount;
	final int cellCount;
	
	//getting row count
	public ExcelUtility(String xlfile,String xlsheet) throws IOException 
	{	
		fis = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fis);
		sheet = wb.getSheet(xlsheet);
		this.rowCount = sheet.getLastRowNum();
		row = sheet.getRow(this.rowCount);
		this.cellCount =row.getLastCellNum();
		wb.close();
		fis.close();
		
		
	}
	
	//getting cell count
	public int getCellCount() throws IOException 
	{
		return cellCount;
	}
	public int getRowCount() throws IOException 
	{
		return rowCount;
	}
	
	
	//get cell data
	public static String getCelldata(String xlfile,String xlsheet, int rowNum,int colnum) throws IOException
	{
		fis = new FileInputStream(xlfile);
		wb = new XSSFWorkbook(fis);
		sheet = wb.getSheet(xlsheet);
		row = sheet.getRow(rowNum);
		cell = row.getCell(colnum);
		String data;
		try
		{
			DataFormatter formatter = new DataFormatter();
			String cellData = formatter.formatCellValue(cell);
			return cellData;
		}
		catch(Exception e)
		{
			data="";
		}wb.close();
		fis.close();
		return data;
	}

	
}
