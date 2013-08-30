package ru.chernobrivenko;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class ReadTemplate {
	
	Excel ex;
	
	List<List<Coord>> coordList;
	
	ReadTemplate(Excel ex)
	{
		this.ex = ex;
	}
	
	public class Coord
	{
		int row, column;
		
		Coord(int row, int column)
		{
			this.row = row;
			this.column = column;
		}
	}
	
	private class Label
	{
		String text;
		int row, column;
		
		Label(String text)
		{
			this.text = text;  
			loadRowColumn();
		}
		
		void loadRowColumn()
		{
			Iterator<Row> it = ex.getRowIterator();
			while(it.hasNext())
			{
				Row r = it.next();
				
				if(r.getCell(0).getCellType() == Cell.CELL_TYPE_STRING 
						&& r.getCell(0).getRichStringCellValue().toString().equals(text))
				{
					row = r.getRowNum();
					column = 0;
				}
			}
		}
	}
	
	public void loadTemplate(String src)
	{
		Element root = getXMLRootElement(src);
		
		Map<String,Label> lables = new HashMap<String, Label>();
		
		NodeList nl = root.getElementsByTagName("label");
		
		for(int i = 0; i < nl.getLength(); i++)
		{
			Element el = (Element) nl.item(i);
			String s = el.getAttribute("name");
			Label t = new Label(el.getAttribute("text"));
			lables.put(s, t);
		}
		
		coordList = new ArrayList<List<Coord>>();
		
		nl = root.getElementsByTagName("row");	
		for(int i = 0; i < nl.getLength(); i++)
		{
			List<Coord> lt = new ArrayList<Coord>();
			
			NodeList nlCell = ((Element) nl.item(i)).getElementsByTagName("cell");
			
			for(int j = 0; j < nlCell.getLength(); j++ )
			{
				Element cell = (Element) nlCell.item(j);
				
				Label label = lables.get(cell.getAttribute("label"));
				
				String rowS = cell.getAttribute("row");
				int row = label.row;
				if(!rowS.equals(""))
					row += Integer.parseInt(rowS);
				
				String columnS = cell.getAttribute("column");
				int column = label.column;
				if(!columnS.equals(""))
					column += Integer.parseInt(columnS);
				
				Coord c = new Coord(row, column);
				
				lt.add(c);
			}
			
			coordList.add(lt);
		}
		
		nl = root.getElementsByTagName("rows");	
		for(int i = 0; i < nl.getLength(); i++)
		{
			Element el = (Element) nl.item(i);
			
			Label labelFrom = lables.get(el.getAttribute("fromLabel"));
			
			String rowFromS = el.getAttribute("fromShift");
			int rowFrom = labelFrom.row;
			if(!rowFromS.equals(""))
				rowFrom += Integer.parseInt(rowFromS);
						
			int columnFrom = labelFrom.column;
			
			Label labelTo = lables.get(el.getAttribute("toLabel"));
			
			String rowToS = el.getAttribute("toShift");
			int rowTo = labelTo.row;
			if(!rowToS.equals(""))
				rowTo += Integer.parseInt(rowToS);
			
			String[] columns = el.getAttribute("columns").split(",");
			
			for(int j = rowFrom; j < rowTo; j++)
			{
				List<Coord> lt = new ArrayList<Coord>();
				
				for(int k = 0; k < columns.length; k++)
					lt.add(new Coord(j, columnFrom + Integer.parseInt(columns[k])));
					
				coordList.add(lt);
				
			}
			
			
		}
	}
		
	
	public Iterator<List<Coord>> getIterator()
	{
		return coordList.iterator();
	}
	
	private Element getXMLRootElement(String filename)
	{
		Element root = null;
		try {
			File fXmlFile = new File(filename);
			DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
			dbFactory.setNamespaceAware(true);
			DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.parse(fXmlFile);	
			
			root = doc.getDocumentElement();
			root.normalize();
		} catch (Exception e) {

			e.printStackTrace();			
		}
		return root;

	}
}
