package org.isp1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadMixedData {
public static void main(String[] args) throws IOException {
	
	File excelLoc = new File("C:\\Users\\USER\\eclipse-workspace\\MavenTest1\\Excel\\zupain.xlsx");
	FileInputStream fIS = new FileInputStream(excelLoc);
	Workbook w = new XSSFWorkbook(fIS);
	Sheet s = w.getSheet("Zupain");
	for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
		Row r = s.getRow(i);
		for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
			Cell c = r.getCell(j);
			int type = c.getCellType();
			if (type==1)
			{
			String sValue = c.getStringCellValue();
			System.out.println(sValue);		
			}
			else if (type==0)
			
			{
				if (DateUtil.isCellDateFormatted(c)) 
				
				{
				 Date dt = c.getDateCellValue();
				 SimpleDateFormat sdf = new SimpleDateFormat("MMM/dd/yyyy");
				 String dValue = sdf.format(dt);
				 System.out.println(dValue);
						
				}
					else {
						double d = c.getNumericCellValue();
						long l = (long)d;
						String pNumber = String.valueOf(l);
						 System.out.println(pNumber);
						
						
					}
							
				}
				
				
				
				
				
			}
			
				
				
				
			}	
			}
			
			
			
		}
		

	
	
	


