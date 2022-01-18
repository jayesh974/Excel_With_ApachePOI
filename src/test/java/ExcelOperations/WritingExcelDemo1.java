package ExcelOperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelDemo1 {
	
	public static void main(String[] args) throws IOException {
		
		Object [][] empdata = 
			{
				{"EmpId","Name","Job"},
				{101,"Jayesh","Engineer"},
				{102,"Sunil","Tester"},
				{103,"Sachine","Databasetester"}
 		    };
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp info");
		
//		 using for loop
		int rows= empdata.length;
		int cols = empdata[1].length;
		
		for(int r=0;r<rows ;r++)
		{
			XSSFRow row = sheet.createRow(r);
			
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell = row.createCell(c);
				
				Object value = empdata[r][c];
				
				if(value instanceof String)
					cell.setCellValue((String) value);
				if(value instanceof Integer)
					cell.setCellValue((Integer) value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean) value);
			}
					    
        }
		
		String filePath = ".\\datafiles\\employee.xlsx" ;
		
		FileOutputStream outstream = new FileOutputStream(filePath);
		workbook.write(outstream);
		
		outstream.close();
		
		System.out.println("employee.xlsx written successfully");
		
		
		
	}

}
