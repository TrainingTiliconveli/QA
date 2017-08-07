package com.veli;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.*;
//import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class ExcelBook 
{
   public static void main(String[] args) throws Exception 
   {
      //Create blank workbook
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      //Create a blank sheet
      XSSFSheet spreadsheet = workbook.createSheet( " Employee Info ");
      //Create row object
      XSSFRow row;
      //This data needs to be written (Object[])
      Map < String, Object[] > empinfo = 
      new TreeMap < String, Object[] >();
      empinfo.put( "1", new Object[] { 
      "EMP ID", "EMP NAME", "DESIGNATION","CONTACT NO" });
      empinfo.put( "2", new Object[] { 
      "tp01", "Gopal", "Technical Manager" ,9856456523L});
      empinfo.put( "3", new Object[] { 
      "tp02", "Manisha", "Proof Reader" ,8956234174L});
      empinfo.put( "4", new Object[] { 
      "tp03", "Masthan", "Technical Writer" });
      empinfo.put( "5", new Object[] { 
      "tp04", "Satish", "Technical Writer" });
      empinfo.put( "6", new Object[] { 
      "tp05", "Krishna", "Technical Writer" });
      //Iterate over data and write to sheet
      Set < String > keyid = empinfo.keySet();
      int rowid = 0;
      for (String key : keyid)
      {
         row = spreadsheet.createRow(rowid++);
         Object [] objectArr = empinfo.get(key);
         int cellid = 0;
         for (Object obj : objectArr)
         {
        	 
            XSSFCell cell = row.createCell(cellid++);
            cell.setCellValue((String)obj);
         }
      }
      //Write the workbook in file system
      FileOutputStream out = new FileOutputStream( 
      new File("Writesheet.xlsx"));
      workbook.write(out);
      FileInputStream fis = new FileInputStream(
    	      new File("WriteSheet.xlsx"));
      Iterator < Row > rowIterator = spreadsheet.iterator();
      while (rowIterator.hasNext()) 
      {
         row = (XSSFRow) rowIterator.next();
         Iterator cells = row.cellIterator();
         while (cells.hasNext()) {
             XSSFCell cell = (XSSFCell) cells.next();

         CellType type = cell.getCellTypeEnum();
         if (type == CellType.STRING) {
             System.out.println("[" + cell.getRowIndex() + ", "
                     + cell.getColumnIndex() + "] = STRING; Value = "
                     + cell.getRichStringCellValue().toString());
         } else if (type == CellType.NUMERIC) {
             System.out.println("[" + cell.getRowIndex() + ", "
                     + cell.getColumnIndex() + "] = NUMERIC; Value = "
                     + cell.getNumericCellValue());
         } else if (type == CellType.BOOLEAN) {
             System.out.println("[" + cell.getRowIndex() + ", "
                     + cell.getColumnIndex() + "] = BOOLEAN; Value = "
                     + cell.getBooleanCellValue());
         } else if (type == CellType.BLANK) {
             System.out.println("[" + cell.getRowIndex() + ", "
                     + cell.getColumnIndex() + "] = BLANK CELL");
         }
     }
 }
         System.out.println();
         fis.close();  
         out.close();
         // workbook.close();
          System.out.println( 
          "Writesheet.xlsx written successfully" );
   }
     
     
   
   }
