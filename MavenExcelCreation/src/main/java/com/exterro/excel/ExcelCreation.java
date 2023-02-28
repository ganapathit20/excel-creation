package com.exterro.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCreation {
	
	public static void getExcel() throws Exception {
		
		XSSFWorkbook workBook = new XSSFWorkbook();

		Sheet sheet = workBook.createSheet("demo");
		 
		Map<String, Object[]> data = new TreeMap<String, Object[]>();

		data.put("1",
		        new Object[] { "ID", "NAME", "LASTNAME" });
		data.put("2",
		        new Object[] { 1, "Pankaj", "Kumar" });
		data.put("3",
		        new Object[] { 2, "Prakashni", "Yadav" });
		data.put("4", new Object[] { 3, "Ayan", "Mondal" });
		data.put("5", new Object[] { 4, "Virat", "kohli" });

		int row = 0;

		Set<String> keyset = data.keySet();

		for (String key : keyset) {

			Row row1 = sheet.createRow(row++);
			
			Object[] objArr = data.get(key);
			
			int col = 0;

			for (Object obj : objArr) {

				Cell cell = row1.createCell(col++);

				if (obj instanceof String)
		            cell.setCellValue((String)obj);

		        else if (obj instanceof Integer)
		            cell.setCellValue((Integer)obj);
			}

		}

		FileOutputStream out = new FileOutputStream(new File("test.xlsx"));

		workBook.write(out);

		out.close();

		System.out.println("done..");
	}

}
