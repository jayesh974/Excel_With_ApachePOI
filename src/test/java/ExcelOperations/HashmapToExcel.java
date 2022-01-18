package ExcelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashmapToExcel {
	
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();		
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		
		Map<String, String> data = new HashMap<String,String>();
		
		data.put("101", "jayesh");
		data.put("102", "sachin");
		data.put("103", "aditya");
		data.put("104", "sunil");
		
		int rowcount = 0;
		for(Map.Entry entry:data.entrySet())
		{
			XSSFRow row = sheet.createRow(rowcount++);
			
			row.createCell(0).setCellValue((String)entry.getKey());
			row.createCell(1).setCellValue((String)entry.getValue());
		
		}
		
		FileOutputStream fos = new FileOutputStream(".\\datafiles\\studet.xlsx");
		workbook.write(fos);
		fos.close();
		
		System.out.println("Excel written successfully......");
		
		
			
		
		
		
		
		
	}

}
