package mypackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Utility {
	
	
	public static HSSFWorkbook filename;
	public static ArrayList<String> getRow(String cellContent,String TestfileName,int sheetNumber)
			throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {
		FileInputStream file = new FileInputStream(new File(TestfileName));
		POIFSFileSystem fs = new POIFSFileSystem(file);
		filename = new HSSFWorkbook(fs);
		HSSFSheet sheet = filename.getSheetAt(sheetNumber);
		InputStream fileIn = new FileInputStream(TestfileName);
		Workbook wb = WorkbookFactory.create(fileIn);
		int flag = 0;
		ArrayList<String> cells = new ArrayList<String>();
		try {
			for (Row row : sheet) {
				for (Cell cell : row) {
					if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
						// if
						// (cell.getRichStringCellValue().getString().trim().equals(cellContent))
						// {
						if (cell.getRichStringCellValue().getString().trim().contains(cellContent)) {
							flag = 1;
							cell = row.getCell(row.getRowNum());
							flag = 1;
							Iterator<Cell> cellItr = row.iterator();
							while (cellItr.hasNext()) {
								cells.add(cellItr.next().toString());
							}
							return cells;
						}
					}

					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						double temp = cell.getNumericCellValue();
						String temps = temp + "";
						if (temps.contains(cellContent) == true) {
							flag = 1;
							cell = row.getCell(row.getRowNum());
							flag = 1;
							Iterator<Cell> cellItr = row.iterator();
							while (cellItr.hasNext()) {
								cells.add(cellItr.next().toString());
							}
							return cells;
						}
					}

					wb.close();
				}
			}
			if (flag == 0) {
				System.out.println("No match found!");
			}
		} catch (Exception E) {
			System.out.println(E);
		}
		return null;
	}


}
