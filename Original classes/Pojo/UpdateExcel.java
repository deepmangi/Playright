package Pojo;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateExcel {

    public static void main(String[] args) {
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
}