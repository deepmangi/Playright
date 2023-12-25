package Utility;

import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class usingutility {
	private static String filePath = "C:\\Users\\deep.mangi\\Desktop\\student2.xlsx";
	private static String sheetName = "Sheet1";
	
	
	@Test
	public static void usingExclUility() throws IOException, InvalidFormatException {
		
		ExcelUtility eUtil = new ExcelUtility(filePath, sheetName);
		
		//getting row count
		int rows = eUtil.getRowCount();
		System.out.println("Total Rows are:- " + rows);
		
		//getting cell count
		int cols = eUtil.getCellCount();
		System.out.println("Total Columns are:-" + cols);
		
		//getting data
		String data = eUtil.getCelldata(filePath, sheetName, 5, 4);
		System.out.println("This is the data:-" + data );
		
//		eUtil.wrieExcel(filePath, sheetName, "Col 5,Row 1");
	}
	
	@Test
	public static void Utility2() {
		
		//This data needs to be written (Object[])
				Map<String, Object[]> data = new TreeMap<String, Object[]>();
				data.put("1", new Object[] {"ID", "FIST NAME", "LAST NAME"});
				data.put("2", new Object[] {1, "Test", "1"});
				data.put("3", new Object[] {2, "Test", "2"});
				data.put("4", new Object[] {3, "Test", "3"});
				data.put("5", new Object[] {4, "Test", "4"});
				data.put("6", new Object[] {4, "Test", "5"});
				
		excelUtilityComplete.Write("New Sheet", filePath, data);
	}
	@Test
	public static void updateutil2() {
				
		excelUtilityComplete.Update(filePath,1,"Test","7");
	}
}
