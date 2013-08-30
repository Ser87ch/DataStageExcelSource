package ru.chernobrivenko;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Properties;


import org.apache.poi.OldFileFormatException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ascentialsoftware.jds.Stage;


public class ReadExcelStage extends Stage {
	private Iterator<Row> rows;

	public void initialize() 
	{ 
		//Java Pack API Logging method to log messages to Director Log.
		info("*****Initializing Excel Application*****"); 

		//Code block to initalize and read user property values provided in the Stage GUI.
		byte[] userProperties = getUserProperties().getBytes();
		InputStream propertyStream = new ByteArrayInputStream(userProperties);
		Properties properties = new Properties();

		try
		{
			properties.load(propertyStream);  //User properties loaded.
		}

		//Log a message if user properties cannot be loaded.
		catch(IOException e)
		{
			info("*****Could not load user properties");
		} 	

		//Read the value of the user property named ExcelFileName
		String propertyValue = properties.getProperty("ExcelFileName" ); 		

		//Code block to create the objects of 
		//the File input stream,workbook and sheet.
		try 
		{
			FileInputStream fis = new FileInputStream(propertyValue);
			Workbook workbook = new XSSFWorkbook(fis);


			//Get the first sheet object
			Sheet sheet = workbook.getSheetAt(0);

			rows = sheet.rowIterator(); 
		}	 
		catch (FileNotFoundException e) 
		{
			//Log a message if the input excel file is not found
			info("The input Excel file is not found"); 
		}
		catch(IOException ioExp)
		{
			//Log the IO Exception
			info("*****IO Exception*****"); 
		}	
		catch(OldFileFormatException UnsupportedExcelFormatException)
		{
			//Log a message if the Excel format is not supported.
			info("*****Cannot read old Excel format*****"); 
		}		
	} 


	public void terminate()
	{
		// Log information message while terminating the Application.	
		info("*****Terminating Excel Application*****");
	}


	//The core processing logic of the Java application.
	public int process() 
	{	 
		//If rows exist
		if(rows.hasNext())								
		{
			int colCount = 0;

			//Create an output row object to hold the row data
			com.ascentialsoftware.jds.Row outputRow	= createOutputRow(); 		

			//Create an Excel row object
			Row hrow = rows.next();	

			//Create a cell iterator for the row
			Iterator<Cell> cells = hrow.cellIterator();	

			//If cells exist
			while (cells.hasNext())						
			{
				Cell hcell = cells.next();

				//Extract cell value
				String cellData = extractCellValue(hcell,hcell.getCellType());	

				//Assign the row data to the outout row object
				outputRow.setValueAsString(colCount,cellData);		
				colCount++;
			}
			info("*****Writing Excel data to Target*****");			
			//Write the row to the output link in the job
			writeRow(outputRow);
			return OUTPUT_STATUS_READY;				
		}
		return OUTPUT_STATUS_END_OF_DATA;
	} 

	protected String extractCellValue(Cell cell, int hcellType)
	{
		String cell_value;

		switch (hcellType) 					
		{
		//If Cell is blank
		case Cell.CELL_TYPE_BLANK:	
			cell_value = "";
			break;

			//If Cell value is boolean
		case Cell.CELL_TYPE_BOOLEAN:
			cell_value = "" + cell.getBooleanCellValue();
			break;

			//If Cell value is string
		case Cell.CELL_TYPE_STRING:
			cell_value = cell.getRichStringCellValue().toString();
			break;

			//Invalid cell
		case Cell.CELL_TYPE_ERROR:		
			cell_value = "ERROR";
			break;

			//If Cell value is numeric
		case Cell.CELL_TYPE_NUMERIC:
			cell_value = Double.toString(cell.getNumericCellValue()); //formatter.formatCellValue(cell, formulaEval);
			break;
			//Default Cell value
		default:							
			cell_value = "DEFAULT_VALUE";
			//cell_value = cell.getRichStringCellValue().toString();
			break;
		}
		return cell_value;
	}	

}
