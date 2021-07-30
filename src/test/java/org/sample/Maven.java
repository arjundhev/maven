package org.sample;

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

public class Maven {
	public static void main(String[] args) throws IOException {
		File file=new File("C:\\Users\\ELCOT\\eclipse-workspace\\Arjunan\\FirstMavenProject\\workbook\\mavenlaunch.xlsx");
		FileInputStream fileInputStream=new FileInputStream(file);
		Workbook book=new XSSFWorkbook(fileInputStream);
		Sheet sheet=book.getSheet("sheet1");
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			Row row=sheet.getRow(i);
			for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
				Cell cell=row.getCell(j);
				int cellType = cell.getCellType();
				
				if (cellType==1) {
					String s = cell.getStringCellValue();
					System.out.println(s);
						
				} else if (DateUtil.isCellDateFormatted(cell)) {
					Date dateCellValue = cell.getDateCellValue();
					SimpleDateFormat format=new SimpleDateFormat("MMM-dd-yyyy");
					String format2 = format.format(dateCellValue);
					System.out.println(format2);
					
				} 
				else {
					
				
				       double numericCellValue = cell.getNumericCellValue();
				     Long l=(long)numericCellValue;
				       System.out.println(l);
				       
				} }}}}

			
					
				
				



				
