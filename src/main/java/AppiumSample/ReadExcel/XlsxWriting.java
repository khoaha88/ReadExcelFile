package AppiumSample.ReadExcel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellAlignment;

public class XlsxWriting {
	static void WriteXLSX(String strCurrentSheet, int rowIndex, int cellIndex, String strStatus, String strLog) throws IOException{
		String excelFileName = "template.xlsx";//name of excel file
//		String sheetName = strCurrentSheet;//name of sheet
		
		FileInputStream fileInput = new FileInputStream(excelFileName);
		
//		@SuppressWarnings("resource")
//		XSSFWorkbook wb = new XSSFWorkbook();
//		XSSFSheet sheet = wb.createSheet(strCurrentSheet) ;
//		XSSFRow row = sheet.createRow((short)rowIndex);
//		XSSFCell cell = row.createCell(cellIndex);
//		cell.setCellValue(strStatus);
		XSSFWorkbook wb =  new XSSFWorkbook(fileInput);
		XSSFSheet sheet = wb.getSheet(strCurrentSheet);
		XSSFCell cellStatus = sheet.getRow(rowIndex).createCell(cellIndex);
		XSSFCell cellLog = sheet.getRow(rowIndex).createCell(cellIndex + 1);
		
		cellStatus.setCellValue(strStatus);
		cellLog.setCellValue(strLog);
		fileInput.close();
		
		FileOutputStream fileOut = new FileOutputStream(excelFileName);
		//write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
}
