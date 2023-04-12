package org.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelData {

	//public static void main(String[] args) throws IOException {
		
		public String ReadValueFromExcel(String path, String SheetName, int rowNum, int cellNum) throws IOException {

			String res = null;

			File file = new File(path);
			FileInputStream fileIn = new FileInputStream(file);
			Workbook workbook = new XSSFWorkbook(fileIn);
			Sheet sheet = workbook.getSheet(SheetName);
			Row row = sheet.getRow(rowNum);
			Cell cell = row.getCell(cellNum);
			CellType type = cell.getCellType();

			switch (type) {

			case STRING:

				res = cell.getStringCellValue();
				break;

			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {

					Date dateCellValue = cell.getDateCellValue();
					SimpleDateFormat dateFormat = new SimpleDateFormat("dd/mm/yy");

					res = dateFormat.format(dateCellValue);

				}

				else {

					double numericCellValue = cell.getNumericCellValue();
					long check = Math.round(numericCellValue);

					if (check == numericCellValue) {

						res = String.valueOf(check);

					} else {

						res = String.valueOf(numericCellValue);

					}

				}
				break;

			default:

				break;

			}
			return res;
		}
	}
	
	
