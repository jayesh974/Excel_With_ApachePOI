package ExcelOperations;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromFormulaCell {
	
	public static void main(String[] args) throws IOException {
		
		FileInputStream filepath = new FileInputStream(".\\datafiles\\readformula.xlsx");
		
		XSSFWorkbook workbook = new XSSFWorkbook(filepath);
		
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();
		
		for(int i=0; i<rows ; i++)
		{
			XSSFRow row = sheet.getRow(i);
			
			for(int c=0; c<cols ;c++)
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
				System.out.print(" |  ");
			}
			System.out.println();
		}
		filepath.close();
		
	}

}
