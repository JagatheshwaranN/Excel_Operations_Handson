package com.jtaf.excel.handson;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
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

	public String fetchCellData(String sheetName, int colNum, int rowNum) {

		try {
			if (rowNum <= 0)
				return "";
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(colNum);
			if (cell == null)
				return "";
			if (cell.getCellType() == CellType.STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA) {
				String cellData = String.valueOf(cell.getNumericCellValue());
				if (DateUtil.isCellDateFormatted(cell)) {
					double d = cell.getNumericCellValue();
					Calendar calendar = Calendar.getInstance();
					calendar.setTime(DateUtil.getJavaDate(d));
					cellData = calendar.get(Calendar.DAY_OF_MONTH) + "/" + calendar.get(Calendar.DAY_OF_MONTH) + 1 + "/"
							+ String.valueOf(calendar.get(Calendar.YEAR)).substring(2);
				}
				return cellData;
			} else if (cell.getCellType() == CellType.BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());

		} catch (Exception ex) {
			ex.printStackTrace();
			return "row " + rowNum + " or column " + colNum + " is doesn't exists in the xlsx workbook";
		}
	}

	public String fetchCellData(String sheetName, String colName, int rowNum) {

		try {

			if (rowNum <= 0)
				return "";
			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName.trim()))
					col_Num = i;
			}
			if (col_Num == -1)
				return "";
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(col_Num);
			if (cell == null)
				return "";
			if (cell.getCellType() == CellType.STRING)
				return cell.getStringCellValue();
			else if (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA) {
				String cellData = String.valueOf(cell.getNumericCellValue());
				if (DateUtil.isCellDateFormatted(cell)) {
					double d = cell.getNumericCellValue();
					Calendar calendar = Calendar.getInstance();
					calendar.setTime(DateUtil.getJavaDate(d));
					cellData = calendar.get(Calendar.DAY_OF_MONTH) + "/" + calendar.get(Calendar.DAY_OF_MONTH) + 1 + "/"
							+ String.valueOf(calendar.get(Calendar.YEAR)).substring(2);
				}
				return cellData;
			} else if (cell.getCellType() == CellType.BLANK)
				return "";
			else
				return String.valueOf(cell.getBooleanCellValue());

		} catch (Exception ex) {
			ex.printStackTrace();
			return "row " + rowNum + " or column " + colName + " is doesn't exists in the xlsx workbook";
		}
	}

}
