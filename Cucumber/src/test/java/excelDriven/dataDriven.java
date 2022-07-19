package excelDriven;

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

public class dataDriven {

	
	public ArrayList<String> getData(String testcaseName) throws IOException
	{

				ArrayList<String> a=new ArrayList<String>();
				
				FileInputStream fs=new FileInputStream("C://Users//002OER744//Documents//Book1.xlsx");
				XSSFWorkbook wb=new XSSFWorkbook(fs);
				
				int sheets=wb.getNumberOfSheets();
				for(int i=0;i<sheets;i++)
				{
					if(wb.getSheetName(i).equalsIgnoreCase("Sheet1"))
							{
					XSSFSheet sheet=wb.getSheetAt(i);
					
					
					 Iterator<Row>  rows= sheet.iterator();
					Row firstrow= rows.next();
					Iterator<Cell> ce=firstrow.cellIterator();
					int k=0;
					int col = 0;
				while(ce.hasNext())
				{
					Cell value=ce.next();
					
					if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
					{
						col=k;
						
					}
					
					k++;
				}
				System.out.println(col);
				
				
				while(rows.hasNext())
				{
					
					Row r=rows.next();
					
					if(r.getCell(col).getStringCellValue().equalsIgnoreCase(testcaseName))
					{
						
						
						
						Iterator<Cell>  cv=r.cellIterator();
						while(cv.hasNext())
						{
						Cell c=	cv.next();
						if(c.getCellTypeEnum()==CellType.STRING)
						{
							
						a.add(c.getStringCellValue());
						}
						else{
							
							a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
						
						}
						}
					}
						
				}		
					
					
							}
				}
				return a;
				
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		
	}

}

