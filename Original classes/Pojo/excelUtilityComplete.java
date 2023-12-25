package Pojo;

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
	public static void Write(String sheet_Name,String path) 
	{
		try (//Blank workbook
				XSSFWorkbook workbook = new XSSFWorkbook()) {
			//Create a blank sheet
			XSSFSheet sheet = workbook.createSheet(sheet_Name);

			//This data needs to be written (Object[])
			Map<String, Object[]> data = new TreeMap<String, Object[]>();
			data.put("1", new Object[] {"ID", "NAME", "LAST NAME"});
			data.put("2", new Object[] {1, "Test", "1"});
			data.put("3", new Object[] {2, "Test", "2"});
			data.put("4", new Object[] {3, "Test", "3"});
			data.put("5", new Object[] {4, "Test", "4"});

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
	public static void Update() {

		XSSFWorkbook workbook=null;
		XSSFSheet sheet;
		try{
			FileInputStream file = new FileInputStream(new File("C:\\Users\\deep.mangi\\Desktop\\Product.xlsx"));

			//Create Workbook instance holding reference to .xlsx file
			workbook = new XSSFWorkbook(file);

			//Get first/desired sheet from the workbook
			//Most of people make mistake by making new sheet by looking in tutorial
			sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());

			Employee ess = new Employee(5,"Test","5");
			//Get the count in sheet
			int rowCount = sheet.getLastRowNum()+1;
			Row empRow = sheet.createRow(rowCount);
			System.out.println();
			Cell c1 = empRow.createCell(0);
			c1.setCellValue(ess.getId());
			Cell c2 = empRow.createCell(1);
			c2.setCellValue(ess.getFirstName());
			Cell c3 = empRow.createCell(2);
			c3.setCellValue(ess.getLastName());
		}
		catch (Exception e) 
		{
			e.printStackTrace();
		}
		try
		{
			//Write the workbook in file system
			FileOutputStream out = new FileOutputStream(new 
					File("C:\\Users\\deep.mangi\\Desktop\\Product.xlsx"));
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
	private String firstName;
	private String lastName;

	public excelUtilityComplete(){}

	public excelUtilityComplete(int id, String firstName, String lastName) {
		super();
		this.id = id;
		this.firstName = firstName;
		this.lastName = lastName;
	}

	public String getFirstName() {
		return firstName;
	}
	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}
	public String getLastName() {
		return lastName;
	}
	public void setLastName(String lastName) {
		this.lastName = lastName;
	}
	public int getId() {
		return id;
	}
	public void setId(int id) {
		this.id = id;
	}
}