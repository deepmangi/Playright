package Pojo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
    import java.util.Set;
    import java.util.TreeMap;        
    import org.apache.poi.ss.usermodel.Cell;
    import org.apache.poi.ss.usermodel.Row;
    import org.apache.poi.xssf.usermodel.XSSFSheet;
    import org.apache.poi.xssf.usermodel.XSSFWorkbook;

    public class ApachePOIExcelWrite {
                public static void main(String[] args) 
                {
                    try (//Blank workbook
					XSSFWorkbook workbook = new XSSFWorkbook()) {
						//Create a blank sheet
						XSSFSheet sheet = workbook.createSheet("Employee Data");

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
						    FileOutputStream out = new FileOutputStream(new File("C:\\Users\\deep.mangi\\Desktop\\Product.xlsx"));
						    workbook.write(out);
						    out.close();
						    System.out.println("Write Successfully.");
						} 
						catch (Exception e) 
						{
						    e.printStackTrace();
						}
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
                }
    }