package com.leaarning.testng;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public static void main(String[] args) {

		try {

			FileInputStream fis = new FileInputStream(new File("C:\\Users\\prash\\Desktop\\Demo.xlsx"));
			

			XSSFWorkbook workbook = new XSSFWorkbook(fis);

			XSSFSheet sheet = workbook.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			System.out.println(rows);

			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row  = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch(cell.getCellType()){
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println(cell.getNumericCellValue() + "\t\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.println(cell.getStringCellValue() + "\t\t");
						break;
					}
				}
				System.out.println("");
			}
			workbook.close();
			fis.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
