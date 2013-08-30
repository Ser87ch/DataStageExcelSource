package ru.chernobrivenko;

import static java.lang.System.out;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Excel {

	private Iterator<Row> rows;
	private Sheet sheet;


	public Excel()
	{

	}


	public void loadExcel(String src) throws FileNotFoundException, IOException
	{
		FileInputStream fis = new FileInputStream(src);
		Workbook workbook = new XSSFWorkbook(fis);


		//Get the first sheet object
		sheet = workbook.getSheetAt(0);

		rows = sheet.rowIterator();	



	}

	public Iterator<Row> getRows()
	{
		return rows;
	}

	private List<String> getStringArrayFromRow()
	{
		if(rows.hasNext())
		{
			List<String> ls = new ArrayList<String>();

			//Create an Excel row object
			Row hrow = rows.next();	

			//Create a cell iterator for the row
			Iterator<Cell> cells = hrow.cellIterator();	

			//If cells exist
			while (cells.hasNext())						
			{
				Cell hcell = cells.next();

				//Extract cell value
				String cellData = extractCellValue(hcell);	

				ls.add(cellData);

			}
			return ls;
		}
		else
			return null;
	}

	public boolean writeOutputRow(com.ascentialsoftware.jds.Row outputRow)
	{
		List<String> ls = getStringArrayFromRow();

		if(ls != null)
		{
			for(int i = 0; i < ls.size(); i++)
			{
				outputRow.setValueAsString(i, ls.get(i));
			}

			return true;
		}
		else
			return false;

	}

	private String extractCellValue(Cell cell)
	{
		String cell_value;

		int type = cell.getCellType();

		if(type == Cell.CELL_TYPE_FORMULA)
			type = cell.getCachedFormulaResultType();

		switch (type) 					
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
			if (DateUtil.isCellDateFormatted(cell))
			{
				Date dt = cell.getDateCellValue();
				if(new SimpleDateFormat("yyyy").format(dt).equals("1899"))
					cell_value = new SimpleDateFormat("HH:mm").format(dt);
				else
					cell_value = new SimpleDateFormat("dd/MM/yyyy").format(dt); 
			}
			else
				cell_value = Double.toString(cell.getNumericCellValue()); 
			break;
			//Default Cell value		
		default:							
			cell_value = "DEFAULT_VALUE";
			//cell_value = cell.getRichStringCellValue().toString();
			break;
		}
		return cell_value;
	}

	public void print()
	{
//		CellReference cellReference = new CellReference("A26"); 
//		Row row = sheet.getRow(cellReference.getRow());
//		Cell cell = row.getCell(cellReference.getCol()); 
		
//		out.println(cell.getRichStringCellValue().toString());

		List<String> ls = getStringArrayFromRow();
		while(ls != null)
		{
			out.println(ls.toString());

			ls = getStringArrayFromRow();
		}
	}

}
