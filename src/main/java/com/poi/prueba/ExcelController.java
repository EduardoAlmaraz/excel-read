package com.poi.prueba;

import java.io.File;
import java.io.FileOutputStream;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class ExcelController {
	
	@Autowired
	private ExcelBusiness excelBusiness;
	
	private static final String FILENAME = "/opt/archivos/excel.xlsx";

	@PostMapping(value="/process-excel")
	@ResponseBody
	public void processExcel(@RequestParam MultipartFile file) {
		
		try(FileOutputStream fos = new FileOutputStream(new File(FILENAME))) {
			fos.write(file.getBytes());
			excelBusiness.processExcel(FILENAME);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
