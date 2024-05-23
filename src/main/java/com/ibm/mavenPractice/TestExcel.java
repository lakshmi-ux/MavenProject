package com.ibm.mavenPractice;

import java.io.IOException;

public class TestExcel {

	public static void main(String[] args) throws IOException {
		String s= ExcelCode.readStringData(0,0);
		System.out.println(s);
		
		double d= ExcelCode.readNumData(2, 2);
		System.out.println(d);
		

	}

}
