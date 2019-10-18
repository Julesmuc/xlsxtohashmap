package com.schroeter.exceltohashmap;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class App 
{
	public static final String SAMPLE_XLSX_FILE_PATH = "/src/main/resources/example.xlsx";

	public static void main(String[] args) throws IOException {
		Workbook workbook = WorkbookFactory.create(new File(System.getProperty("user.dir")+SAMPLE_XLSX_FILE_PATH));
		Sheet sheet = workbook.getSheetAt(0);
		Map<Integer, Map<String, String>> dataMap = new HashMap<>();
		DataFormatter dataFormatter = new DataFormatter();
		List<String> header = new ArrayList<>();
		sheet.forEach(row -> {
			Map<String, String> map = new HashMap<>();
			row.forEach(cell -> {
				String cellValue = dataFormatter.formatCellValue(cell);
				if (row.getRowNum() == 0) {
					header.add(cellValue);
				} else {
					map.put(header.get(cell.getColumnIndex()), cellValue);
					dataMap.put(row.getRowNum(), map);
				}
			});
		});
		
		dataMap.entrySet().forEach(entry -> 
			System.out.println((entry.getKey() + " " + entry.getValue().toString())));
		workbook.close();
	}

}
