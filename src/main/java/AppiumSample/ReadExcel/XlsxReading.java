package AppiumSample.ReadExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;

public class XlsxReading {
	@SuppressWarnings("deprecation")
	public static void ReadXLSX() throws IOException{
		FileInputStream input_document = new FileInputStream(new File("template.xlsx")); //Read XLSX document - Office 2007, 2010 format     
        @SuppressWarnings("resource")
		XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document); //Read the Excel Workbook in a instance object    
        XSSFSheet my_worksheet = my_xlsx_workbook.getSheetAt(0); //This will read the sheet for us into another object
        Iterator<Row> rowIterator = my_worksheet.iterator(); // Create iterator object
        List<String> lstRowValue;
        List<String> lsGetResult;
        Row row;
        Cell cell;
//        XlsxWriting xlsxWriting = new XlsxWriting();
        while(rowIterator.hasNext()) {
       	 	lstRowValue = new ArrayList<String>();
            row = rowIterator.next(); //Read Rows from Excel document 
            
            if (row.getRowNum() < 1) continue;
            
            for (int i =0; i<row.getLastCellNum(); i++) {
            	cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
           	 	lstRowValue.add(cell.toString());
            }
            lsGetResult = DoStep(lstRowValue);
            XlsxWriting.WriteXLSX("Sheet1", 1, Constant.colStatusIndex, lsGetResult.get(0), lsGetResult.get(1));
            System.out.println(Arrays.toString(lstRowValue.toArray())); // To iterate over to the next row
        }
        input_document.close(); //Close the XLS file opened for printing
	}
	
	public static List<String> DoStep(List<String> ActionsOfStep){
		List<String> lsReturn = new ArrayList<String>();
		String strAction = ActionsOfStep.get(2).toUpperCase();
		String strValue = ActionsOfStep.get(3);
		switch(strAction) {
		case "OPEN":
			String strException = openBrowser(strValue);
			if (strException == "" || strException == null){
				lsReturn.add("PASSED");
				lsReturn.add("");
			}
			else {
				lsReturn.add("FAILED");
				lsReturn.add(strException);
			}
			break;
		default:
			break;
		}
		return lsReturn;
	}
	
//	static 
	public static String openBrowser(String url){
		WebDriver driver;
//		List<String> lsReturn = new ArrayList<String>();
		String log;
		try{
			System.setProperty("webdriver.chrome.driver",".//references//chromedriver.exe"); 
			driver=new ChromeDriver();
			driver.manage().window().maximize();
			driver.get(url);
			driver.close();
			log = "";
//			lsReturn.add("true");
//			lsReturn.add("");			
		}
		catch(Exception ex){
			log = ex.toString();
//			lsReturn.add("false");
//			lsReturn.add(ex.toString());
		}
		return log;
	}
}
