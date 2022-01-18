package ExcelOperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingPasswordProtectedFile {
	
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		String path = ".\\datafiles\\customers.xlsx";
		FileInputStream fis = new FileInputStream(path);
		String password = "test123";
		
//		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis,password);
		XSSFSheet sheet = workbook.getSheetAt(0);	
		
//	**********	Read data from excel sheet using for loop*************
/*					
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(0).getLastCellNum();
		System.out.println(rows);   //starting from 0
		System.out.println(cols);   // starting from 1
		
		for(int r=0;r<=rows;r++)
		{
			XSSFRow row = sheet.getRow(r);
			
			for(int c=0; c<cols;c++)
			{
				XSSFCell cell = row.getCell(c);
				
				switch(cell.getCellType())
				{
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case FORMULA :
					System.out.print(cell.getNumericCellValue());
				}
				System.out.print(" | ");
			}
			System.out.println();
		}
	*/	
		
//		***********read data from excel sheet using iterator********
		
		Iterator<Row> iterator = sheet.iterator();
		
		while(iterator.hasNext())
		{
			Row nextrow = iterator.next();
			
			Iterator<Cell> celliterator = nextrow.cellIterator();
			
			while(celliterator.hasNext())
			{
				Cell cell = celliterator.next();
				
				switch(cell.getCellType())
				{
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue());
					break;
				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue());
					break;
				case FORMULA :
					System.out.print(cell.getNumericCellValue());
				}
				System.out.print(" | ");				
			}
			System.out.println();
		}
	
	}

}
