package org.isp1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MavenExcel {

	public static void main(String[] args) throws IOException {
		
		File excelLoc = new File("C:\\Users\\USER\\eclipse-workspace\\MavenTest1\\Excel\\Usha Medico (1).xlsx");
		FileInputStream fIn = new FileInputStream(excelLoc);
		Workbook w = new XSSFWorkbook(fIn);
		
		Sheet s = w.getSheet("Sheet1");
		Row r = s.getRow(5);
		Cell c = r.getCell(1);
		System.out.println(c);
		int rows = s.getPhysicalNumberOfRows();
		System.out.println(rows);
		int cells = r.getPhysicalNumberOfCells();
		System.out.println(cells);
		

		
	}
		
}
