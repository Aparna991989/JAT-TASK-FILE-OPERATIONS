package Utils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

		 import java.io.FileOutputStream;
		 import java.io.IOException;
		 
		 public class WriteExcelDemo {

		
		     public static void main(String[] args) throws IOException {
		         // Create a new workbook
		         Workbook workbook = new XSSFWorkbook();

		         // Create a sheet
		         Sheet sheet = workbook.createSheet("sheet1");

		         // Create data to write
		         Object Data[][] = {
		             {"Name", "Age", "Email"},
		             {"John Doe", "30", "john@test.com"},
		             {"Jane Doe", "28", "jane@test.com"},
		             {"Bob Smith", "35", "jacky@example.com"},
		             {"Swapnil" ,"37", "swapnil@example.com"}
		             };

		         int rows  = Data.length;
		         int cols= Data[0].length;
		         
		        for(int r =0; r < rows; r++) {
		        	XSSFRow row =(XSSFRow) sheet.createRow(r);
		        	 for(int c =0; c < cols; c++) {
		        		 XSSFCell cell = row.createCell(c);
		        		 Object value = Data[r][c];
		        		 if(value instanceof String)
		        			 cell.setCellValue((String)value);
		        		 if(value instanceof Integer)
		        			 cell.setCellValue((Integer)value);
		        		 if(value instanceof Boolean)
		        			 cell.setCellValue((Boolean)value);
		        	 }
		        }
		         
		     String filePath= ".\\excelsheet\\sheet1.xlsx";
		     FileOutputStream fis = new FileOutputStream(filePath);
		     workbook.write(fis);
		     fis.close();
		         
		     
		 }

		 }
