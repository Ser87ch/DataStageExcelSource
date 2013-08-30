package ru.chernobrivenko;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import com.ascentialsoftware.jds.Stage;


public class ReadExcel extends Stage {
	private Iterator<Row> rows;
	private HSSFFormulaEvaluator formulaEval;
	public void initialize() 
	{ 
		//Java Pack API Logging method to log messages to Director Log.
		info("*****Initializing Excel Application*****"); 

		//Code block to initalize and read user property values provided in the Stage GUI.
		byte[] userProperties 		= getUserProperties().getBytes();
		InputStream propertyStream 	= new ByteArrayInputStream(userProperties);
		Properties properties 		= new Properties();

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
		String propertyValue 	= properties.getProperty("ExcelFileName" ); 		

		//Code block to create the objects of 
		//the File input stream,workbook and sheet.
		try 
		{
			FileInputStream fis		=new FileInputStream(propertyValue);
			HSSFWorkbook workbook			=new HSSFWorkbook(fis);
			formulaEval			=new HSSFFormulaEvaluator(workbook);

			//Get the first sheet object
			HSSFSheet sheet = workbook.getSheetAt(0);

			rows				=sheet.rowIterator(); 
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
		catch(org.apache.poi.hssf.OldExcelFormatException UnsupportedExcelFormatException)
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
			int colCount		= 0;

			//Create an output row object to hold the row data
			com.ascentialsoftware.jds.Row outputRow 		= createOutputRow(); 		

			//Create an Excel row object
			HSSFRow hrow 		= (HSSFRow) rows.next();	

			//Create a cell iterator for the row
			Iterator<Cell> cells 	= hrow.cellIterator();	

			//If cells exist
			while (cells.hasNext())						
			{
				HSSFCell hcell 	= (HSSFCell) cells.next();

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

	protected String extractCellValue(HSSFCell cell, int hcellType)
	{
		String cell_value;
		DataFormatter formatter = new DataFormatter();
		switch (hcellType) 					
		{
		//If Cell is blank
		case HSSFCell.CELL_TYPE_BLANK:	
			cell_value = "";
			break;

			//If Cell value is boolean
		case HSSFCell.CELL_TYPE_BOOLEAN:
			cell_value = "" + cell.getBooleanCellValue();
			break;

			//If Cell value is string
		case HSSFCell.CELL_TYPE_STRING:
			cell_value = cell.getRichStringCellValue().toString();
			break;

			//Invalid cell
		case HSSFCell.CELL_TYPE_ERROR:		
			cell_value = "ERROR";
			break;

			//If Cell value is numeric
		case HSSFCell.CELL_TYPE_NUMERIC:
			cell_value = formatter.formatCellValue(cell, formulaEval);
			break;

			//Default Cell value
		default:							
			cell_value = "DEFAULT_VALUE";
			break;
		}
		return cell_value;
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub

	}

}
