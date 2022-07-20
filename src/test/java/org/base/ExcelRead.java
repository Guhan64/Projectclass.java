package org.base;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Table.Cell;

public class ExcelRead 
{
	public static void main(String[] args) throws IOException 
	{
		File loc=new File ("C:\\Users\\HP\\eclipse-workspace\\MavenProject\\src\\test\\resources\\Excel.xlsx");
		FileInputStream fi=new FileInputStream (loc);	
		
		Workbook w=new XSSFWorkbook(fi);
		
		Sheet s=w.getSheet("Sheet1");
		
		Row r=s.getRow(2);
		
	    Cell c =r.getCell(0);
		System.out.println(c);
		
		int rowCount =s.getPhysicalNumberOfRows();
		System.out.println(rowCount);
		
		int cellCount =r.getPhysicalNumberOfCells();
		System.out.println(cellCount);
		
		for(int i=0;i<s.getPhysicalNumberOfRows();i++)
		{Row row=s.getRow(i);
		for(int j=0;j<r.getPhysicalNumberOfCells();j++)
		{Cell cell2=r.getCell(j);
		System.out.println(cell2);
		}
		}
		
	}

}
