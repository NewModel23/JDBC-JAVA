import java.awt.Font;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainActivity {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		Workbook workbook = new XSSFWorkbook();
		@SuppressWarnings("unused")
		org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("HojaDePrueba");
		
		org.apache.poi.ss.usermodel.Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		
		// Create a Row
		Row headerRow = sheet.createRow(0);
		
		Cell cell = headerRow.createCell(1);
		cell.setCellStyle(headerCellStyle);
		
		cell.setCellValue("Prueba funcionó");
		
		int rowNum = 1;

		FileOutputStream fileOut = new FileOutputStream("Pruebas.xlsx");
		try {
			workbook.write(fileOut);
			
			System.out.print("Archivo Creado con éxito!");
			
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.print(e.getMessage());
		}
		finally {
		fileOut.close();
		}
		
		
	}

}
