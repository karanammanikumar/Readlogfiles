import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class project2 {
     static XSSFWorkbook workbook = new XSSFWorkbook(); 
     static XSSFSheet spreadsheet = workbook.createSheet("Generated "+LocalDate.now());
     static XSSFRow row;
     static XSSFCell cell;
      static CellStyle style = workbook.createCellStyle();

      public static void main( String[] args ) throws IOException {
      File file =new File("D:/Public/file/MOL.ksh_EXPORT_MOL_01t.log");
      int i ;
      style.setDataFormat(workbook.createDataFormat().getFormat("hh:mm:ss"));
      String s1 = "EnteringprocessNTIMOL";
      String s2 = "Valid documents :";
      String s3 = "Entering copyDVD1";
      String s4 = "Entering processDirCreateForCDorDVDWrite";

      Scanner in = null;
      row = spreadsheet.createRow(0);
      cell = row.createCell(1);
      cell.setCellValue("HU_HU");
      cell = row.createCell(2);
      cell.setCellValue("RO_RO");
      
      row = spreadsheet.createRow(1);
      cell = row.createCell(0);
      cell.setCellValue("Start of NTI");

      try {
    	  i=1;
          in = new Scanner(file);
          while(in.hasNext())
          {
              String line=in.nextLine();
                  if(line.contains(s1))
                  {
                	  String str[] = line.split(",");
                	  System.out.println(str[0]+"------"+s1);
                      cell = row.createCell(i);
                      cell.setCellStyle(style);
                      cell.setCellValue(str[0]);
                      i++;
                  }
            
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }
      
      row = spreadsheet.createRow(2);
      cell = row.createCell(0);
      cell.setCellValue("End of NTI");
      try {
    	  i=1;
          in = new Scanner(file);
          while(in.hasNext())
          {
              String line=in.nextLine();
                  if(line.contains(s2))
                  {
                	  String str[] = line.split(",");
                	  System.out.println(str[0]+"------"+s2);
                      cell = row.createCell(i);
                      cell.setCellStyle(style);
                      cell.setCellValue(str[0]);
                      i++;
                  }
            
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }

      row = spreadsheet.createRow(3);
      cell = row.createCell(0);
      cell.setCellValue("Difference");
      for(i=1;i<=44;i++)
      {
    	  
           cell = row.createCell(i);
           cell.setCellFormula("B3-B2");
                 
      }
      
      
      row = spreadsheet.createRow(5);
      cell = row.createCell(0);
      cell.setCellValue("Start of TM");
      try {
    	  i=1;
          in = new Scanner(file);
          while(in.hasNext())
          {
              String line=in.nextLine();
                  if(line.contains(s2))
                  {
                	  String str[] = line.split(",");
                	  System.out.println(str[0]+"------"+s2);
                      cell = row.createCell(i);
                      cell.setCellStyle(style);
                      cell.setCellValue(str[0]);
                      i++;
                  }
            
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }
      
      row = spreadsheet.createRow(6);
      cell = row.createCell(0);
      cell.setCellValue("End of TM");
      try {
    	  i=1;
          in = new Scanner(file);
          while(in.hasNext())
          {
              String line=in.nextLine();
                  if(line.contains(s3))
                  {
                	  String str[] = line.split(",");
                	  System.out.println(str[0]+"------"+s3);
                      cell = row.createCell(i);
                      cell.setCellStyle(style);
                      cell.setCellValue(str[0]);
                      i++;
                  }
            
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }
      
      row = spreadsheet.createRow(9);
      cell = row.createCell(0);
      cell.setCellValue("start of MR");
      try {
    	  i=1;
          in = new Scanner(file);
          while(in.hasNext())
          {
              String line=in.nextLine();
                  if(line.contains(s3))
                  {
                	  String str[] = line.split(",");
                	  System.out.println(str[0]+"------"+s3);
                      cell = row.createCell(i);
                      cell.setCellStyle(style);
                      cell.setCellValue(str[0]);
                      i++;
                  }
            
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }
      
      
      row = spreadsheet.createRow(10);
      cell = row.createCell(0);
      cell.setCellValue("End of MR");

      try {
    	  i=1;
          in = new Scanner(file);
          while(in.hasNext())
          {
              String line=in.nextLine();
                  if(line.contains(s4))
                  {
                	  String str[] = line.split(",");
                	  System.out.println(str[0]+"------"+s4);
                      cell = row.createCell(i);
                      cell.setCellStyle(style);
                      cell.setCellValue(str[0]);
                      i++;
                  }
            
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }


      System.out.println(spreadsheet.getLastRowNum());
      FileOutputStream out = null;
			try {
				String exepath = "D://exceldata";
								String bufpath = exepath.concat(".xlsx");
								out = new FileOutputStream (new File(bufpath));
								workbook.write(out);
								System.out.println("Executed successfully 2");
//								SendFromYahoo.send();
								out.close();
						
			} catch (Exception e) {
				e.printStackTrace();
			}

		}

	}
