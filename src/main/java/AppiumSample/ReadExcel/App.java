package AppiumSample.ReadExcel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	
	public static void main(String[] args) throws IOException { 
		XlsxReading.ReadXLSX();
//		XlsxWriting.WriteXLSX("Sheet1", 0, 0, "False");
			
//		csvReading.readCSVFile();
//		ReadXLSX();
//		for(int i = 0; i< GetSheetsName().size(); i++)
//			System.out.println(GetSheetsName().get(i));
//		WriteXLSX();
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
		XSSFCell cellB1 = row1.createCell(1);
		cellA1.setCellValue("first row");
		cellB1.setCellValue("cell 2");
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

	
	@SuppressWarnings("deprecation")	
	private static void ReadXLSX() throws IOException{
		 FileInputStream input_document = new FileInputStream(new File("template.xlsx")); //Read XLSX document - Office 2007, 2010 format     
         @SuppressWarnings("resource")
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
             }
             System.out.println(Arrays.toString(lstRowValue.toArray())); // To iterate over to the next row
         }
         input_document.close(); //Close the XLS file opened for printing
	}
}