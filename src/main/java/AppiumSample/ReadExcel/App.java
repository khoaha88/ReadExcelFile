package AppiumSample.ReadExcel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.record.CellValueRecordInterface;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	
	public static void main(String[] args) throws IOException { 
		CSVReading csvReading = new CSVReading();
//		csvReading.readCSVFile();
		ReadXLSX();
	}
	

	private static void WriteXLSX() throws IOException{
		String excelFileName = "writexlsx.xlsx";//name of excel file
		String sheetName = "Sheet1";//name of sheet
		FileOutputStream fileOut = new FileOutputStream(excelFileName);
		
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName) ;
		XSSFRow row1 = sheet.createRow((short)0);
		@SuppressWarnings("deprecation")
		XSSFCell cellA1 = row1.createCell(0);
		
		cellA1.setCellValue("first row");

		//write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		
//		//iterating r number of rows
//		for (int r=0;r < 5; r++ )
//		{
//			XSSFRow row = sheet.createRow(r);
//
//			//iterating c number of columns
//			for (int c=0;c < 5; c++ )
//			{
//				XSSFCell cell = row.createCell(c);
//	
//				cell.setCellValue("Cell "+r+" "+c);
//				
//				XSSFCellStyle cellStyle = wb.createCellStyle();
////		        cellStyle.setFillForegroundColor(XSSFColor.toXSSFColor(color).GOLD.index);
//
//		        cell.setCellStyle(cellStyle);
//			}
//		}
	}
	
	private static List<String> GetSheetsName() throws IOException{
		FileInputStream input_document = new FileInputStream(new File("template.xlsx"));
		 XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document);
		 Iterator<Sheet> sheetIterator;
		 List<String> sheetNames = new ArrayList<String>();
		 for (int i=0; i<my_xlsx_workbook.getNumberOfSheets(); i++) {
		    sheetNames.add( my_xlsx_workbook.getSheetName(i) );
		}
		return sheetNames;
	}
	
	private static void ReadFile() throws IOException{
		FileInputStream input_document = new FileInputStream(new File("template.xlsx"));
		 XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document);
		 XSSFSheet my_worksheet = my_xlsx_workbook.getSheet("Sheet1");
		 Iterator<Row> rowIterator = my_worksheet.iterator();
		 
//		 my_xlsx_workbook.gets
		 XSSFRow row1 = my_worksheet.getRow(0);
//		 System.out.print(row1.getLastCellNum());
		 while(rowIterator.hasNext()){
			 Row row = rowIterator.next();
			 Iterator<Cell> cellIterator = row.cellIterator();
			 while(cellIterator.hasNext()){
				 Cell cell = cellIterator.next();
				 System.out.println(cell.getStringCellValue() + "\t\t");
			 }
			 System.out.println("");
		 }
	}
	private static void ReadXLSX() throws IOException{
		 FileInputStream input_document = new FileInputStream(new File("template.xlsx")); //Read XLSX document - Office 2007, 2010 format     
         XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document); //Read the Excel Workbook in a instance object    
         XSSFSheet my_worksheet = my_xlsx_workbook.getSheetAt(0); //This will read the sheet for us into another object
         Iterator<Row> rowIterator = my_worksheet.iterator(); // Create iterator object
         List<String> lstRowValue;
         Row row;
         Cell cell;
         while(rowIterator.hasNext()) {
        	 lstRowValue = new ArrayList<String>();
             row = rowIterator.next(); //Read Rows from Excel document       
             for (int i =0; i<row.getLastCellNum(); i++){
            	 cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
            	 lstRowValue.add(cell.toString());
//            	 System.out.print(cell.toString() + ", ");
             }
             System.out.println(Arrays.toString(lstRowValue.toArray())); // To iterate over to the next row
         }
         input_document.close(); //Close the XLS file opened for printing
	}
}