package com.jtaf.excel.scenarios;

import com.jtaf.excel.handson.ExcelReader;

public class GetExcelSheetSuiteRunnable {

	public static void main(String[] args) {

		System.out.println(checkSuiteRunnable());
	}

	public static boolean checkSuiteRunnable() {

		ExcelReader reader = new ExcelReader(System.getProperty("user.dir") + "/src/test/resources/TestWorkBook.xlsx");
		String sheetName = "TestSuite";
		String suiteCol = "Suite";
		String runModeCol = "RunMode";
		int rows = reader.getSheetRowCount(sheetName);
		for (int row = 2; row <= rows; row++) {

			String data = reader.fetchCellData(sheetName, suiteCol, row);
			if (data.equalsIgnoreCase("BankManagerSuite")) {
				String runMode = reader.fetchCellData(sheetName, runModeCol, row);
				if (runMode.equalsIgnoreCase("Y"))
					return true;
				else
					return false;
			}
		}
		return false;
	}
}
