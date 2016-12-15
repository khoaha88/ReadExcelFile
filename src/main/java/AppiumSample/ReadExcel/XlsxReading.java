package AppiumSample.ReadExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsxReading {
	public static void ReadXLSX() throws IOException{
		FileInputStream input_document = new FileInputStream(new File("template.xlsx")); //Read XLSX document - Office 2007, 2010 format     
        XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document); //Read the Excel Workbook in a instance object    
        XSSFSheet my_worksheet = my_xlsx_workbook.getSheetAt(0); //This will read the sheet for us into another object
        Iterator<Row> rowIterator = my_worksheet.iterator(); // Create iterator object
        String [] arrRowData;
        while(rowIterator.hasNext()) {
        	
//                Row row = rowIterator.next(); //Read Rows from Excel document       
//                Iterator<Cell> cellIterator = row.cellIterator();//Read every column for every row that is READ
//                        while(cellIterator.hasNext()) {
//                                Cell cell = cellIterator.next(); //Fetch CELL
//                                switch(cell.getCellType()) { //Identify CELL type
//                                case Cell.CELL_TYPE_NUMERIC:
//                                        System.out.print(cell.getNumericCellValue() + "\t\t"); //print numeric value
//                                        break;
//                                case Cell.CELL_TYPE_STRING:
//                                        System.out.print(cell.getStringCellValue() + "\t\t"); //print string value
//                                        break;
//                                }
//                        }
        	System.out.println(""); // To iterate over to the next row
        }
        input_document.close(); //Close the XLS file opened for printing
	}
}
