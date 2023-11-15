package com.roytuts.apache.poi.excel.write.data.generic.way;

import java.util.ArrayList;
import java.util.List;

public class ExcelFileWriterApp {

	public static void main(String[] args) {
		List<Person> persons = new ArrayList<>();

		Person p1 = new Person("A", "a@roytuts.com", "Kolkata");
		Person p2 = new Person("B", "b@roytuts.com", "Mumbai");
		Person p3 = new Person("C", "c@roytuts.com", "Delhi");
		Person p4 = new Person("D", "d@roytuts.com", "Chennai");
		Person p5 = new Person("E", "e@roytuts.com", "Bangalore");
		Person p6 = new Person("F", "f@roytuts.com", "Hyderabad");

		persons.add(p1);
		persons.add(p2);
		persons.add(p3);
		persons.add(p4);
		persons.add(p5);
		persons.add(p6);

		ExcelFileWriter.writeToExcel("excel-person.xlsx", persons);
	}

}
