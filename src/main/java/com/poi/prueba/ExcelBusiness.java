package com.poi.prueba;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Service;

@Service
public class ExcelBusiness {
	
//	HSSFRow row;
//	HSSFCell cell;
	StringBuffer sb = new StringBuffer();
	Row row;
	Cell cell;
	

	public void processExcel(String filename) throws Exception {
		
		FileInputStream fis = new FileInputStream(new File(filename));
		
//		HSSFWorkbook workbook = new HSSFWorkbook(fis);
		
		Workbook workbook = WorkbookFactory.create(fis);
		Sheet sheet = workbook.getSheetAt(0);
		
		workbook.close();
		int total = sheet.getLastRowNum();
		for (int i = 0; i < total; i++) {
			row = sheet.getRow(i);
			generateCsv("/opt/archivos/prueba.csv", row);
		}
	}
	
	private void generateCsv(String filename, Row row) {
		sb.setLength(0);
		for (int j = 0; j < 6; j++) {
			cell = row.getCell(j);
			cell.setCellType(CellType.STRING);
			String content = cell.getStringCellValue();
			sb.append(content);
			if(j < 5)
				sb.append(",");
			else
				sb.append("\n");
		}
		System.out.println(sb.toString());
		File file = new File(filename);
		if(!file.exists()) {
			try {
				file.createNewFile();
				FileWriter fileWriter = new FileWriter(file, true);
				BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);
			    bufferedWriter.write(sb.toString());
				fileWriter.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
	}
}
