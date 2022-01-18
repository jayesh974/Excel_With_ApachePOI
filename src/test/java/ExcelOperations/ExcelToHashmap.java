package ExcelOperations;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToHashmap {
	
	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream(".\\datafiles\\studet.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		HashMap<String,String> data = new HashMap<String,String>();	
		
		int rows = sheet.getLastRowNum();
		
		for(int r=0; r<=rows ; r++)
		{
			String key = sheet.getRow(r).getCell(0).getStringCellValue();
			String value = sheet.getRow(r).getCell(1).getStringCellValue();
			
			data.put(key, value);			
		}
		
		System.out.println(data);
		
		for(Map.Entry entry:data.entrySet())
		{
			System.out.println(entry.getKey()+"   "+entry.getValue());
		}
		
	
		
	}

}
