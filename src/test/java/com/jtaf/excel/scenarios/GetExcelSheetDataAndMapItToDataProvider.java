package com.jtaf.excel.scenarios;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.jtaf.excel.handson.ExcelReader;

public class GetExcelSheetDataAndMapItToDataProvider {

	@Test(dataProvider = "getData")
	public void dataFromProvider1(String customer, String currency, String message) {

		System.out.println(customer + "---" + currency + "---" + message);
	}

	@DataProvider(name = "getData")
	public Object[][] getDataFromExcel() {
		Object[][] data = null;
		ExcelReader reader = new ExcelReader(System.getProperty("user.dir") + "/src/test/resources/TestWorkBook.xlsx");
		String sheetName = "TestDataSet";
		int rows = reader.getSheetRowCount(sheetName);
		String testName = "OpenAccountTest";
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

		data = new Object[testDataRows][testDataCols];

		// Printing the actual test data of each test
		for (int rowNum = testDataStartRowNum; rowNum < (testDataStartRowNum + testDataRows); rowNum++) {

			for (int colNum = 0; colNum < testDataCols; colNum++) {

				data[rowNum - testDataStartRowNum][colNum] = reader.fetchCellData(sheetName, colNum, rowNum);
			}
		}
		return data;
	}
}
