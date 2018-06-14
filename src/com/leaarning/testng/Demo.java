package com.leaarning.testng;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {
	public void demo() throws IOException {
		FileInputStream inputStream = new FileInputStream("C:\\Users\\prash\\Desktop\\Demo.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sh = workbook.getSheetAt(0);

		// To get the number of rows present in sheet
		int totalNoOfRows = sh.getLastRowNum()+1;

		// To get the number of columns present in sheets
		
		XSSFRow row = sh.getRow(0);
		int totalNoOfCols = row.getLastCellNum();

		for (int col = 1; col < totalNoOfCols; col++) {

			for (int r = 1; r < totalNoOfRows; r++) {
				String data = sh.getRow(r).getCell(col).getStringCellValue();
				System.out.println(data);
			}
			System.out.println();
		}
		workbook.close();
	}

	public static void main(String args[]) throws IOException {
		Demo DT = new Demo();
		DT.demo();
	}
}

