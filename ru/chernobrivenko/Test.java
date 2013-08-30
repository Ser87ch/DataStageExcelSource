package ru.chernobrivenko;

import java.io.FileNotFoundException;
import java.io.IOException;

public class Test {

	
	public static void main(String[] args) {
		Excel ex = new Excel();
		try {
			ex.loadExcel("C:\\Users\\chernobrivenko\\Downloads\\Example DDR1.xlsx","template.xml");
				
			ex.print();
		} catch (FileNotFoundException e) {
			
			e.printStackTrace();
		} catch (IOException e) {
			
			e.printStackTrace();
		}

	}

}
