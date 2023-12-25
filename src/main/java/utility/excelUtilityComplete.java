package utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;        
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelUtilityComplete {
	
	//this map stores the data which is to be stored in excel
	static Map<String, Object[]> data = new TreeMap<String, Object[]>();

	public static void Write(String sheet_Name,String path,Map<String, Object[]> data) 
	{
		try (
			//Blank workbook
			XSSFWorkbook workbook = new XSSFWorkbook()) {
			
			//Create a blank sheet
			XSSFSheet sheet = workbook.createSheet(sheet_Name);

			//Iterate over data and write to sheet
			Set<String> keyset = data.keySet();
			int rownum = 0;
			for (String key : keyset)
			{
				Row row = sheet.createRow(rownum++);
				Object [] objArr = data.get(key);
				int cellnum = 0;
				for (Object obj : objArr)
				{
					Cell cell = row.createCell(cellnum++);
					if(obj instanceof String)
						cell.setCellValue((String)obj);
					else if(obj instanceof Integer)
						cell.setCellValue((Integer)obj);
				}
			}
			try
			{
				//Write the workbook in file system
				FileOutputStream out = new FileOutputStream(new File(path));
				workbook.write(out);
				out.close();
				System.out.println("******File Write Successfull.******");
			} 
			catch (Exception e) 
			{
				e.printStackTrace();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public static void Update(String path,int ID,String KEY,String VALUE) {

		XSSFWorkbook workbook=null;
		XSSFSheet sheet;
		try{
			FileInputStream file = new FileInputStream(new File(path));

			//Create Workbook instance holding reference to .xlsx file
			workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			//Most of people make mistake by making new sheet by looking in tutorial
			sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());

			excelUtilityComplete Up = new excelUtilityComplete(ID,KEY,VALUE);
			//Get the count in sheet
			int rowCount = sheet.getLastRowNum()+1;
			Row empRow = sheet.createRow(rowCount);
			System.out.println();
			Cell c1 = empRow.createCell(0);
			c1.setCellValue(Up.getId());
			Cell c2 = empRow.createCell(1);
			c2.setCellValue(Up.getFirstName());
			Cell c3 = empRow.createCell(2);
			c3.setCellValue(Up.getLastName());
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		try
		{
			//Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new 
					File(path));
			workbook.write(out);
			out.close();
			System.out.println("Update Successfully");
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
	
	//employee code
	private int id;
	private String col1;
	private String col2;

	public excelUtilityComplete(){}

	public excelUtilityComplete(int id, String first, String second) {
		super();
		this.id = id;
		this.col1 = first;
		this.col2 = second;
	}

	public String getFirstName() {
		return col1;
	}
	public void setFirstName(String first) {
		this.col1 = first;
	}
	public String getLastName() {
		return col2;
	}
	public void setLastName(String second) {
		this.col2 = second;
	}
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
}