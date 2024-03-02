package com.hybrid;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//      HOW TO READ A SPECIFIC CELL VALUE ?

public class HybridFwork 
{
	static XSSFWorkbook xssfworkbook;
	static XSSFSheet sheet;
	public static void main(String[] args) 
	{
		
		try {
			File file = new File("C:\\Users\\HP\\Documents\\RegTestData.xlsx");
			FileInputStream inputstream = new FileInputStream(file);
			xssfworkbook = new XSSFWorkbook(inputstream);
		} 
		catch (IOException e) 
		{
			System.out.println("File not found, Please check the path again: " +e);
		}
		
		sheet = xssfworkbook.getSheet("Data");
		
		XSSFRow row = sheet.getRow(1);
		XSSFCell cell = row.getCell(1);
		
		String cellvalue = cell.getStringCellValue();
		
		System.out.println(cellvalue);
		
	}

}
