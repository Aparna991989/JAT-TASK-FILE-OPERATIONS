package Utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public static void main(String[] args) {
		getRowCount();
		getCellDataString(0,0);
		getCellDataNumeric(1,1);
		
	}
	
	public static void getRowCount() {
		try {
			String projectPath= System.getProperty("user.dir");
			workbook =new XSSFWorkbook(projectPath+"\\excelsheet\\Task.xlsx");
	sheet = workbook.getSheet("Sheet1");
	int rowCount = sheet.getPhysicalNumberOfRows();
	System.out.println("Number of Rows" +  rowCount);
	}
		catch(IOException e) {
			e.printStackTrace();
		}
		}

	public  static void getCellDataString(int rowNum, int colNum) {
		 try {
			 String projectPath= System.getProperty("user.dir");
				workbook =new XSSFWorkbook(projectPath+"\\excelsheet\\Task.xlsx");
				sheet = workbook.getSheet("sheet1");
				String cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
				System.out.println(cellData);
		 } catch(Exception e) {
				e.printStackTrace();
		 }
	}
	
 public static void getCellDataNumeric(int rowNum, int colNum) {
	 try {
		 String projectPath= System.getProperty("user.dir");
			workbook =new XSSFWorkbook(projectPath+"\\excelsheet\\Task.xlsx");
			sheet = workbook.getSheet("sheet1");
			
			
			int cellData = (int) sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
			System.out.println(cellData);
	 } catch(Exception e) {
			e.printStackTrace();
	 }
	 }
		
}
	


