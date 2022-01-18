package ExcelOperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelDemo2 {

	public static void main(String[] args) throws IOException {

		ArrayList<Object[]> empdata = new ArrayList();

		empdata.add(new Object[] { "EmpId", "Name", "Job" });
		empdata.add(new Object[] { 101, "Jayesh", "Engineer" });
		empdata.add(new Object[] { 102, "Sunil", "Tester" });
		empdata.add(new Object[] { 103, "Sachine", "Databasetester" });

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp info");

//		using for each loop

		int rowcount = 0;

		for (Object[] emp : empdata) {
			XSSFRow row = sheet.createRow(rowcount++);

			int cellcount = 0;
			for (Object value : emp) {
				XSSFCell cell = row.createCell(cellcount++);

				if (value instanceof String)
					if (value instanceof String)
						cell.setCellValue((String) value);
				if (value instanceof Integer)
					cell.setCellValue((Integer) value);
				if (value instanceof Boolean)
					cell.setCellValue((Boolean) value);

			}

		}

		String filePath = ".\\datafiles\\employee2.xlsx";

		FileOutputStream outstream = new FileOutputStream(filePath);
		workbook.write(outstream);

		outstream.close();

		System.out.println("employee2.xlsx written successfully");

	}

}
