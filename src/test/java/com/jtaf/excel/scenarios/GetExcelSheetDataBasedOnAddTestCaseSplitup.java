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
			if (testCaseName.equalsIgnoreCase(testName))
				break;
		}
		System.out.println("Test case starts from row num " + testCaseRowNum);

		// Checking total number of test data rows for each test
		int testDataStartRowNum = testCaseRowNum + 2;
		int testDataRows = 0;
		while (!reader.fetchCellData(sheetName, 0, testDataStartRowNum + testDataRows).equalsIgnoreCase("")) {
			testDataRows++;
		}
		System.out.println("Total rows of test data are " + testDataRows);

		// Checking total number of test data columns for each
		int testDataStartColumNum = testCaseRowNum + 1;
		int testDataCols = 0;
		while (!reader.fetchCellData(sheetName, testDataCols, testDataStartColumNum).equalsIgnoreCase("")) {
			testDataCols++;
		}
		System.out.println("Total columns of test data are " + testDataCols);

		// Printing the actual test data of each test
		for (int rowNum = testDataStartRowNum; rowNum < (testDataStartRowNum + testDataRows); rowNum++) {

			for (int colNum = 0; colNum < testDataCols; colNum++) {

				System.out.println(reader.fetchCellData(sheetName, colNum, rowNum));
			}
		}
	}
}
