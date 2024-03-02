package com.hybrid;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadEntireXL 
{
	static XSSFWorkbook wbook;
	static XSSFSheet sheet;

	public static void main(String[] args) 
	{
		try 
		{
			//Create an object of File class to open xlsx file
			File file = new File("C:\\Users\\HP\\Documents\\RegTestData.xlsx");
			
			//Create an object of FileInputStream class to read excel file
			FileInputStream fis = new FileInputStream(file);
			
			//Creating workbook instance that refers to .xlsx file
			wbook = new XSSFWorkbook(fis);
		} 
		catch (IOException e) 
		{
			System.out.println("Fine not found, Please check the file path again!: "+e);
		}
		
		//creating a Sheet object
		sheet = wbook.getSheet("Data");
		
		// RowCount = LastRowNumber -First Row Number
		
		/*To read the complete data from Excel, you can iterate over each cell of the row, 
		 * present in the sheet. For iterating, you need the total number of rows and cells 
		 * present in the sheet. Additionally, we can obtain the number of rows from the sheet,
		 * which is basically the total number of rows that have data present in the sheet 
		 * by using the calculation -*/
		
		int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
		
		System.out.println("No of Rows presen is: "+rowCount);		
		
		/*Once you get the row, you can iterate over the cells present in the row by using the 
		 * total number of cells, that we can calculate using getLastCellNum() method:*/
		
//		int cellCount = sheet.getRow(4).getLastCellNum();
//		System.out.println("No of Cells present is: "+cellCount);		

		

		for(int i = 0; i<rowCount; i++)
		{
			int cellCount = sheet.getRow(i).getLastCellNum();
			
			System.out.println("Cells present in Row "+i +"is : "+cellCount);
			System.out.println("Row "+i+" Data is: ");
			
			for(int j = 0; j<cellCount; j++)
			{
				System.out.print(sheet.getRow(i).getCell(j).getStringCellValue()+",");
			}
			System.out.println();
			System.out.println();
			
		}
		
		

	}

}
