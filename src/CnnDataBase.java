import com.ibm.as400.access.AS400JDBCDriver;
import java.sql.*;  

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


public class CnnDataBase {

	private static Connection Connection = null;
	static ResultSet rs=null;
	static String username = "";
	static String password = "";
	static String dbUrl = "jdbc:as400://Server";
	
	static String DRIVER = "com.ibm.as400.access.AS400JDBCDriver";
	
	
	public static void main(String[] args ) throws SQLException, ClassNotFoundException, IOException {
		
		Class.forName(DRIVER);
		
		

		GetQuery("Select * from SomeTable");
		
		
		SaveResults();
		
		Connection.close();
	  }
	
	
	public static Connection GetConnection() throws SQLException {
	
		System.out.println("**** Loaded the JDBC driver");
		
		try {
			Connection = DriverManager.getConnection(dbUrl, username, password);
			
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			System.out.print(e.getMessage());
		}
		return Connection;

	}
	
	
	public static void GetQuery(String consulta) throws SQLException {
		
		GetConnection();
		
		Statement stmt = Connection.createStatement();
		rs = stmt.executeQuery(consulta);
		
	}
	
	
	public static void SaveResults() throws IOException, SQLException {
		
		
		 String[] columns = { "Agencia", "Año", "Mes"};
		
		
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
		
		
		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellStyle(headerCellStyle);
		    cell.setCellValue(columns[i]);

		}
		
		int rowNum = 1;
		while(rs.next()) {
			
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(rs.getInt("FIFNIDCIAU"));
			row.createCell(1).setCellValue(rs.getInt("FIFNYEAR"));
			row.createCell(2).setCellValue(rs.getInt("FIFNMONTH"));
			

			rs.next();
		}


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
	

