package com.leaarning.testng;

import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public FileInputStream fis = null;
	public XSSFWorkbook workbook = null;
	public XSSFSheet sheet = null;
	public XSSFRow row = null;
	public XSSFCell cell = null;
	int colCount;
	int rowCount;

	String filePath;

	public ReadData(String filePath) throws Exception {
		this.filePath = filePath;
		fis = new FileInputStream(filePath);
		workbook = new XSSFWorkbook(fis);
		fis.close();
	}

	public int getRowCount(String sheetName) {
		sheet = workbook.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum() + 1;
		return rowCount;

	}

	public int getColumnCount(String sheetName) {
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		return colCount;
	}

	public String getCellData(String sheetName, String colName) {
		StringBuffer buffer = new StringBuffer();
		try {
			int col_Num = -1;
			sheet = workbook.getSheet(sheetName);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num = i;
			}
			for (int j = 1; j <= sheet.getLastRowNum(); j++) {

				row = sheet.getRow(j);
				cell = row.getCell(col_Num);

				if (cell.getCellTypeEnum() == CellType.STRING)
					colName = cell.getStringCellValue();

				else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
					colName = String.valueOf(cell.getNumericCellValue());

				}
//				System.out.println(colName);
				buffer = buffer.append(colName).append("\n");
			}
		} catch (Exception e) {
			e.printStackTrace();
			//return "sheet " + sheetName + " or column " + colName + " does not exist  in Excel";
		}

		return buffer.toString();

	}

	public String getRowData(int rowNum) {
		try {

			sheet = workbook.getSheet("sheet1");
			row = sheet.getRow(rowNum);
			for (int y = 0; y < row.getLastCellNum(); y++) {
				row = sheet.getRow(rowNum);
				cell = row.getCell(y);
				if (cell.getCellTypeEnum() == CellType.STRING)
					System.out.println(cell.getStringCellValue());
				else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
					return String.valueOf(cell.getNumericCellValue());
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "";
	}
}
