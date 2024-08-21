package excelFileHandling;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class ExcelFiles {

	@Test
	public void rc() throws Throwable {

		Workbook wb = WorkbookFactory.create(new FileInputStream("D:\\XSSF.xlsx"));

		Sheet ws = wb.getSheet("Login");
		System.out.println(ws.getLastRowNum());
	}

	@Test
	public void cc() throws Throwable {

		Workbook wb = WorkbookFactory.create(new FileInputStream("D:\\XSSF.xlsx"));

		Sheet ws = wb.getSheet("Login");
		Row row = ws.getRow(0);
		System.out.println(row.getLastCellNum());
	}

	@Test
	public void getData() throws Throwable {

		Workbook wb = WorkbookFactory.create(new FileInputStream("D:\\XSSF.xlsx"));

		Sheet ws = wb.getSheet("Login");
		Row row = ws.getRow(0);

		int rc = ws.getLastRowNum();
		int cc = row.getLastCellNum();

		String[][] data = new String[rc][cc];
		for (int i = 0; i < rc; i++) {
			for (int j = 0; j < cc; j++) {

				if (ws.getRow(i).getCell(j).getCellType() == CellType.NUMERIC) {

					int celldata = (int) ws.getRow(i + 1).getCell(j).getNumericCellValue();
					data[i][j] = String.valueOf(celldata);

				} else {

					data[i][j] = ws.getRow(i + 1).getCell(j).getStringCellValue();

				}
				System.out.print(data[i][j] + " ");

			}
			System.out.println();
		}

	}
	
	@Test
	void write() throws IOException {
		
		FileInputStream fi = new FileInputStream("D://XSSF.xlsx");
		Workbook workbook = WorkbookFactory.create(fi);
		Sheet worksheet = workbook.getSheetAt(0);
		
		for (int i = 1; i <= worksheet.getLastRowNum(); i++) { 

			worksheet.getRow(i).createCell(2).setCellValue("Pass");
			
			Font font = workbook.createFont();
			font.setColor(IndexedColors.GREEN.index);
			font.setBold(true);
			
			CellStyle style = workbook.createCellStyle();
			style.setFont(font);
			
			worksheet.getRow(i).getCell(2).setCellStyle(style);
			
		}
		
		FileOutputStream fo = new FileOutputStream("./Result.xlsx");
		
		workbook.write(fo);
		
		
	}

}
