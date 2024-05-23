package com.ibm.mavenPractice;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCode {
	public static FileInputStream fs;
	public static XSSFWorkbook wb;
	public static XSSFSheet s;
	
//string datas excel sheetil ullathe read cheyan ayite oru method create cheyanam
	public static String readStringData(int i ,int j) throws IOException { //create an excel
		fs= new FileInputStream("D:\\mavenexcels\\Book2.xlsx"); //ee path il kidakunna xlsx file
		//iney java byte ileku convert cheyum, ennite "fs" enna variable ilote store aayi
		
		wb= new XSSFWorkbook(fs); //enni aa fs iney xlsxf workbook ilote store aakum
		//interface aanu enni
		s= wb.getSheet("Sheet1"); //sheet name 
		Row r= s.getRow(i); //row number
		Cell c= r.getCell(j); //cell number
		return c.getStringCellValue();// ee method name inbuild ayi varunathe aanu
	}
	public static double readNumData(int i, int j) throws IOException { //double koduthathe bez int koduthalum eathu numeric 
		//value koduthalum java il  double ayitey eduku
		fs= new FileInputStream("D:\\mavenexcels\\Book2.xlsx"); 
		wb= new XSSFWorkbook(fs); 
		s= wb.getSheet("Sheet1"); 
		Row r= s.getRow(i); 
		Cell c= r.getCell(j); 
		return c.getNumericCellValue();
		
	}
}
