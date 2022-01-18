package ExcelOperations;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormattingCellColour {
	
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet = workbook.createSheet("Sheet");
		
		XSSFRow row = sheet.createRow(1);
		
		XSSFCellStyle style = workbook.createCellStyle();
		style.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN1.getIndex());
		style.setFillPattern(FillPatternType.FINE_DOTS);
		
		XSSFCell cell = row.createCell(1);
		cell.setCellValue("Welcome");
		cell.setCellStyle(style);
				
		FileOutputStream fos = new FileOutputStream(".\\datafiles\\stills.xlsx");
		workbook.write(fos);

		fos.close();
				
	}

}
