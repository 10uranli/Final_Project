import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Write_Excel{
	static String a;
	static double b;
	static boolean c;
	public static void read_Write() throws IOException{
		File myFile = new File("test.xlsx"); 
		FileInputStream fis = new FileInputStream(myFile); //veri al, oku vs de kullanılır.
		// Finds the workbook instance for XLSX file 
		XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
		// Return first sheet from the XLSX workbook 
		XSSFSheet mySheet = myWorkBook.getSheetAt(0);
		// Get iterator to all the rows in current sheet 
		Iterator<Row> rowIterator = mySheet.iterator();//iterator ile collection üzerinde dolaşabiliriz. 
		// Traversing over each row of XLSX file 
		while (rowIterator.hasNext()) { 
			Row row = rowIterator.next(); 
		// For each row, iterate through each columns 
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) { 
			Cell cell = cellIterator.next(); 
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:  a = cell.getStringCellValue(); break; 
		case Cell.CELL_TYPE_NUMERIC: b = cell.getNumericCellValue(); break; 
		case Cell.CELL_TYPE_BOOLEAN: c = cell.getBooleanCellValue(); break; default : } 
		//System.out.println("helloo:" + a);
		write_Excel(a,b,c);
		} 
		
		
		}
		}
	
	
	
	public static void write_Excel(String x,double y, boolean c) throws IOException{
		
		System.out.println("heşşeee"+x);
		
		/*devamm*/
		HSSFWorkbook workbook = new HSSFWorkbook();

		HSSFSheet sheet = workbook.createSheet("Sample Sheet");
		

		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("1", new Object[] {"Ad", "Soyad", "Sınav Salonu"});

		data.put("2", new Object[] {1d, x, 1500000d});

		data.put("3", new Object[] {2d, x, 800000d});

		data.put("4", new Object[] {3d, x, 700000d});
		

	
		


		Set<String> keyset = data.keySet();

		int rownum = 0;

		for (String key : keyset) {

		    Row row = sheet.createRow(rownum++);

		    Object [] objArr = data.get(key);

		    int cellnum = 0;

		    for (Object obj : objArr) {

		        Cell cell = row.createCell(cellnum++);

		        if(obj instanceof Date) 

		            cell.setCellValue((Date)obj);

		        else if(obj instanceof Boolean)

		            cell.setCellValue((Boolean)obj);

		        else if(obj instanceof String)

		            cell.setCellValue((String)obj);

		        else if(obj instanceof Double)

		            cell.setCellValue((Double)obj);

		    }

		}


		try {

		    FileOutputStream out = 

		            new FileOutputStream(new File("test2.xls"));

		    workbook.write(out);

		    out.close();

		    System.out.println("Excel yazıldı..");

		     

		} catch (FileNotFoundException e) {

		    e.printStackTrace();

		} catch (IOException e) {

		    e.printStackTrace();

		}
		
		

		}
	}
