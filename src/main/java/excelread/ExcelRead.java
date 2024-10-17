package excelread;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead 
{
	
	XSSFSheet sheet;
	ExcelRead() throws IOException
	{
		
		FileInputStream inputFile = new FileInputStream("F:\\Obsqura-Selenium\\Reading document for maven.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(inputFile);   //HSSFWorkbook ---if file is type .xls
		
		sheet = workbook.getSheet("Sheet1");
		
	}
	
	public String readExcelData(int i, int j) 
	{
		
		XSSFRow row = sheet.getRow(i);
		Cell cell = row.getCell(j);
		
		CellType type = cell.getCellType();
		
		switch(type)
		{
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());					
		}
		return cell.getStringCellValue();
		
	}
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		ExcelRead ex = new ExcelRead();
		int r = ex.sheet.getLastRowNum();
		ex.sheet.getRow(r).getLastCellNum();
		System.out.println("Excel Data.....");
		for(int i=0; i<ex.sheet.getLastRowNum()+1; i++)
		{
			for(int j=0; j<=ex.sheet.getRow(i).getLastCellNum()-1; j++)
			{
				String s = ex.readExcelData(i,j);
				System.out.print(s + " ");
			}
			System.out.println();
		}
		
	}

}
