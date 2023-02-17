package com.jtaf.excel.handson;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public String path;
	public FileOutputStream fileOutputStream = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;
	private int index;

	public ExcelReader(String path) {
		this.path = path;
		try (FileInputStream fileInputStream = new FileInputStream(path)) {
			workbook = new XSSFWorkbook(fileInputStream);
			sheet = workbook.getSheetAt(0);
		} catch (IOException ex) {
			ex.printStackTrace();
		}
	}

	// Check whether the given sheet is available in the workbook or not.
	public boolean isSheetAvailable(String sheetName) {

		index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1)
				return false;
			else
				return true;
		} else
			return true;
	}

	// Get the total rows count from the given sheet.
	public int getSheetRowCount(String sheetName) {

		if (!isSheetAvailable(sheetName))
			return -1;
		index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			int number = sheet.getLastRowNum() + 1;
			return number;
		}
	}

	// Get the total columns count from the given sheet
	public int getSheetColumnCount(String sheetName) {

		if (!isSheetAvailable(sheetName))
			return -1;
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		if (row == null)
			return -1;
		return row.getLastCellNum();
	}

}
