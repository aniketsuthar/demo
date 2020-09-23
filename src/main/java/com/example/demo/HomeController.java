package com.example.demo;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

@RestController
public class HomeController {

	@PostMapping("/postexcel")
	public void convertToJson(@RequestBody FileInfo fileInfo) throws InvocationTargetException {

		JsonObject sheetsJsonObject = new JsonObject();
		Workbook workbook = null;

		try {
			workbook = new XSSFWorkbook(fileInfo.getFile_path() + fileInfo.getFile_name());
		} catch (IOException e) {
			e.printStackTrace();
		}

		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

			JsonArray sheetArray = new JsonArray();
			ArrayList<String> columnNames = new ArrayList<String>();
			Sheet sheet = workbook.getSheetAt(i);
			Iterator<Row> sheetIterator = sheet.iterator();

			while (sheetIterator.hasNext()) {

				Row currentRow = sheetIterator.next();
				JsonObject jsonObject = new JsonObject();

				if (currentRow.getRowNum() != 0) {
					for (int j = 0; j < currentRow.getLastCellNum(); j++) {

						if (currentRow.getCell(j) != null) {
							if (currentRow.getCell(j).getCellType() == CellType.STRING) {
								jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getStringCellValue());
							} else if (currentRow.getCell(j).getCellType() == CellType.NUMERIC) {
								jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getNumericCellValue());
							} else if (currentRow.getCell(j).getCellType() == CellType.BOOLEAN) {
								jsonObject.addProperty(columnNames.get(j), currentRow.getCell(j).getBooleanCellValue());
							} else if (currentRow.getCell(j).getCellType() == CellType.BLANK) {
								jsonObject.addProperty(columnNames.get(j), "");
							}
						} else {
							jsonObject.addProperty(columnNames.get(j), "");
						}

					}

					sheetArray.add(jsonObject);

				} else {
					// store column names
					for (int k = 0; k < currentRow.getPhysicalNumberOfCells(); k++) {
						columnNames.add(currentRow.getCell(k).getStringCellValue());
					}
				}

			}

			sheetsJsonObject.add(workbook.getSheetName(i), sheetArray);
			System.out.println("File Name: " + fileInfo.getFile_name());
			String[] array = fileInfo.getFile_name().split(".xlsx");
			for (String a : array)
				System.out.println("Array " + a);
			writeStringToFile(sheetsJsonObject.toString(),
					fileInfo.getResult_directory() + workbook.getSheetName(i) + ".json");
		}
	}
	/***
	 * writes json String to the destination folder in .json format for each sheet in workbook
	 * @param data
	 * @param fileName/path
	 */
	private static void writeStringToFile(String data, String fileName)

	{
		try {

			// Get the output file absolute path.
			String filePath = fileName;

			// Create File, FileWriter and BufferedWriter object.
			File file = new File(filePath);

			FileWriter fw = new FileWriter(file);

			BufferedWriter buffWriter = new BufferedWriter(fw);

			// Write string data to the output file, flush and close the buffered writer
			// object.
			buffWriter.write(data);

			buffWriter.flush();

			buffWriter.close();

			System.out.println(filePath + " has been created.");

		} catch (IOException ex) {
			System.err.println(ex.getMessage());
		}
	}

	@GetMapping("/hello")
	public String hello() {
		return "Hello";
	}
	
	@GetMapping("/helloworld")
	@ResponseBody
	public String helloWorld() {
		return "Hello World";
	}
}
