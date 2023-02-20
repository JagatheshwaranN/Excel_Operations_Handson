package com.jtaf.excel.scenarios;

import com.jtaf.excel.handson.ExcelReader;

public class GetExcelSheetRowAndColumnCount {

	public static void main(String[] args) {

		ExcelReader reader = new ExcelReader(System.getProperty("user.dir") + "/src/test/resources/TestWorkBook.xlsx");
		String sheetName = "TestData";
		System.out.println("Excel Sheet Row Count " + reader.getSheetRowCount(sheetName));
		System.out.println("Excel Sheet Col Count " + reader.getSheetColumnCount(sheetName));
	}
}
