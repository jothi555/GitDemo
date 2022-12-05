import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Datadriven {
//Identify the Test cases column
	//identify the purchase test case and read all the columns
	
	public ArrayList<String> getData(String testcaseName) throws IOException
	{
ArrayList<String> a=new ArrayList<String>();
		
		FileInputStream fis=new FileInputStream("C:\\Users\\aramacha\\Javaprojects-Jo\\Demodata\\TestData1.xlsx");
		
	XSSFWorkbook workbook=new XSSFWorkbook(fis);
	int sheets=workbook.getNumberOfSheets();
	
	for (int i=0;i<sheets; i++)
	{
		if (workbook.getSheetName(i).equalsIgnoreCase("Sample1"))
		{
	XSSFSheet sheet=	workbook.getSheetAt(i);
	Iterator<Row>  rows=sheet.rowIterator();
	Row firstrow=rows.next();
	Iterator<Cell> cell=firstrow.iterator();
	int k=0;
	int column=0;
	while(cell.hasNext())
	{
		Cell value=cell.next();
		if (value.getStringCellValue().equalsIgnoreCase("Testcases"))
		{
			//desired column
			column=k;
		}
		k++;
	}
	System.out.println(column);
	
	while(rows.hasNext())
	{
		Row r=rows.next();
		if(r.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName))
		{
			//after u grab the "Purchase" test case, pull all the data in to the rows
			Iterator<Cell> cv=r.cellIterator();
			
			while (cv.hasNext())
			{
				Cell c=cv.next();
				if (c.getCellType()==CellType.STRING)
				{
					a.add(cv.next().getStringCellValue());	
				}
				else 
				{
					
				a.add(NumberToTextConverter.toText(c.getNumericCellValue()));	
				}
			//	System.out.println(cv.next().getStringCellValue());
				
			}
		}
		
		}
	}
		}
	return a;
	}	
	
	public  static void main(String[] args) throws IOException
	{
			
	}
	
}