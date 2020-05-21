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

public class ExcelFile {
	
	public ArrayList<String> getdata(String testcaseName) throws IOException
	{
				ArrayList<String> a = new ArrayList<String>();
		// TODO Auto-generated method stub
				FileInputStream fis = new FileInputStream("D:\\Rohit_Folder\\MAXIMUS\\docs\\demodata.xlsx");//FileInputStream to get data from the location
				XSSFWorkbook workbook = new XSSFWorkbook(fis);//Create XSSFWORKBOOK to keep the data of fis
				
				int sheets = workbook.getNumberOfSheets();//Check Total Number of sheets
				
				for (int i = 0; i < sheets; i++) 
				{
					if (workbook.getSheetName(i).equalsIgnoreCase("TestData"))//Check for the required Sheet name 
					{
						XSSFSheet sheet =  workbook.getSheetAt(i);//GET THE DESIRED SHEET index number
						
					 	Iterator<Row> rows = sheet.iterator();//CREATE A ITERATOR
					 	Row firstrow = rows.next();//ITERATE THROUGH ROWS 
					 	
					 	Iterator<Cell> ce= firstrow.cellIterator();//AFTER ITERATING THORUGH ROWS NOW ITERATE THROUGH CELLS
					 	int k=0;
					 	int coloumn = 0;
					 	while(ce.hasNext())//check wether there are more cells
					 	{
					 		Cell value = ce.next();//if there than move to next
					 		if (value.getStringCellValue().equalsIgnoreCase("TestCase"))//grab value of each cell and check 
					 		{
								coloumn = k;
							}
					 		k++;
					 		
					 		
					 	}
					 	while(rows.hasNext())
					 	{
					 		Row r = rows.next();
					 		if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcaseName))
					 		{
					 			Iterator <Cell> cv= r.cellIterator();
					 			while(cv.hasNext())
					 			{
					 				Cell c = cv.next();//Store cell value in C
					 				if (c.getCellTypeEnum()==CellType.STRING) //Check weather it is STRING
					 				{
					 					a.add(c.getStringCellValue());//Than Store it in a
										
									}
					 				else//is a numeric value than
					 				{
					 					a.add(NumberToTextConverter.toText(c.getNumericCellValue()));//Use NUMBER_TO_TEXT_CONVERTER and grap the NUMBERIC VALUE
					 				}
					 			}
					 			
					 		}
					 	}
					 	
					 	
					 	
						
					}
						
					
				}
				return a;

	}

	
}
