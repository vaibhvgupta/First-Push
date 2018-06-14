package com.leaarning.testng;

public class ReadUtil {

	public static void main(String[] args) {
		try {
			ReadData rd = new ReadData("C:\\Users\\prash\\Desktop\\Demo.xlsx");
			int colCount = rd.getColumnCount("sheetdata");
			System.out.println("Total Colums: " + colCount);
			System.out.println("Colums Value: " + "\n" + rd.getCellData("sheet1", "user"));
			System.out.println(rd.getRowData(4));
		} catch (Exception exception) {
			exception.printStackTrace();
		}
	}
}
