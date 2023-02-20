package com.jtaf.excel.scenarios;

import com.jtaf.excel.handson.ExcelReader;

public class GetExcelSheetDataBasedOnAddTestCaseSplitup {

	public static void main(String[] args) {

		ExcelReader reader = new ExcelReader(System.getProperty("user.dir") + "/src/test/resources/TestWorkBook.xlsx");
		String sheetName = "TestDataSet";
		int rows = reader.getSheetRowCount(sheetName);
		String testName = "AddCustomerTest";
		int testCaseRowNum = 1;
		for (testCaseRowNum = 1; testCaseRowNum <= rows; testCaseRowNum++) {

			String testCaseName = reader.fetchCellData(sheetName, 0, testCaseRowNum);
			System.out.println("testCaseName "+testCaseName);
			System.out.println("testName "+testName);
			if (testCaseName.equalsIgnoreCase(testName))
				break;
		}
		System.out.println("Test case starts from row num " + testCaseRowNum);
	}
}
