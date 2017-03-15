package fr.jk.excelmanipulation.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePOIExcelReader {
	
	private static final String FILE_NAME = "C:/tmp/villes.xlsx";

	public static void main(String[] args) {
			
		try {
			
			FileInputStream cityFile = new FileInputStream(new File(FILE_NAME));
			Workbook workBook = new XSSFWorkbook(cityFile);
			Sheet citySheet = workBook.getSheetAt(0);
			Iterator<Row> rowIterator = citySheet.iterator(); 
			
			while (rowIterator.hasNext()) {
				
				Row currRow = rowIterator.next();
				Iterator<Cell> cellIterator = currRow.iterator();
				
				while (cellIterator.hasNext()) {
					Cell currCell = (Cell) cellIterator.next();
					if (currCell.getCellTypeEnum() == CellType.STRING) {
						System.out.println(currCell.getStringCellValue() + "--");
					} else if (currCell.getCellTypeEnum() == CellType.NUMERIC) {
						System.out.println(currCell.getNumericCellValue() + "--");
					}
				}
				
				System.out.println();
				}
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
	}
}
			
			

