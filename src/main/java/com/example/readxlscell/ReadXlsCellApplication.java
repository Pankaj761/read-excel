package com.example.readxlscell;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

@SpringBootApplication
public class ReadXlsCellApplication {

	public static void main(String[] args) throws IOException {
		SpringApplication.run(ReadXlsCellApplication.class, args);
		InputStream inputStream= new FileInputStream("F:\\data.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(inputStream);
		XSSFCell cell = wb.getSheetAt(0).getRow(4).getCell(2);
		System.out.println(cell.toString());
	}

}
