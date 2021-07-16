package com.lokesh.Excel_To_DB;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class App
{
	public static void main(String[] args)
	{
		try
		{
			String name;
			double salary;
			String desig;
			List<Employee> emp_list = new ArrayList<>();

			File file = new File("D:\\emp_data.xlsx");   //creating a new file instance
			FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file
//creating Workbook instance that refers to .xlsx file  
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
			Iterator<Row> itr = sheet.iterator();    //iterating over excel file
			itr.next();

			while (itr.hasNext())
			{
				Employee emp = new Employee();
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
				while (cellIterator.hasNext())
				{
					Cell cell = cellIterator.next();

					name = cell.getStringCellValue();
					//System.out.println("Name :" + name);
					emp.setName(name);
					cell = cellIterator.next();

					salary = cell.getNumericCellValue();
					//System.out.println("salary :" + salary);
					emp.setSalary(salary);
					cell = cellIterator.next();

					desig = cell.getStringCellValue();
					//System.out.println("desig :" + desig);
					emp.setDesignation(desig);
				}
				System.out.println("");
				emp_list.add(emp);
				emp = null;
			}

			Iterator itr2=emp_list.iterator();
			while(itr2.hasNext())
			{
				Employee obj = (Employee) itr2.next();
				System.out.println("Name :" + obj.getName());

				System.out.println("salary :" + obj.getSalary());

				System.out.println("Desig :" + obj.getDesignation());

				System.out.println(" \n - \n ");
			}

			System.out.println("");
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}


	}
}  