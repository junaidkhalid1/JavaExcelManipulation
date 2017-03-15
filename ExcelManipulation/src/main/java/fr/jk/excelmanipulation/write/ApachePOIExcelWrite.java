package fr.jk.excelmanipulation.write;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePOIExcelWrite {

	private static final String FILE_NAME = "C:/tmp/villes.xlsx";
	
	public static void main(String[] args) {
		
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("Île-de-France");
		
		Object[][] cities = {
				{"Villes" , "Région", "Numéro"},
				{"Paris", "Île-de-France", 75},
				{"Val-de-Marne", "Île-de-France", 94},
				{"Seine Saint-Denis", "Île-de-France", 93},
				{"Hauts-de-Seine", "Île-de-France", 92},
				{"Essone", "Île-de-France", 91},
				{"Seine-et-Marne", "Île-de-France", 77},
				{"Val d'Oise", "Île-de-France", 95},
				{"Yvelines", "Île-de-France", 78},
		};
		
		int rowNum = 0;
		System.out.println("Creating");
		
		for (Object[] city : cities) {
			Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (Object field : city) {
				Cell cell = row.createCell(colNum++);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
			}
		}
		
		try {
			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workBook.write(outputStream);
			workBook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		System.out.println("Created");
	}
}