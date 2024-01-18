package guru;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadUsingIterator {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
FileInputStream fs= new FileInputStream("C:\\Users\\vinishka pranav\\Desktop\\Book1.xlsx");

XSSFWorkbook wrk=new XSSFWorkbook(fs);
XSSFSheet sht= wrk.getSheet("Sheet1");

int row= sht.getLastRowNum();
int col=sht.getRow(1).getLastCellNum();

Iterator iterator= sht.iterator();

while (iterator.hasNext()) {
	XSSFRow rows= (XSSFRow) iterator.next();
	
	Iterator cellit= rows.cellIterator();
	
	while (cellit.hasNext()) {
		
		XSSFCell cells= (XSSFCell) cellit.next();
		
		switch(cells.getCellType()) {
		
		case STRING:
			System.out.print(cells.getStringCellValue()+" ");
			break;
		case NUMERIC:
			System.out.print(cells.getNumericCellValue()+" ");
			break;
		case BOOLEAN:
			System.out.print(cells.getBooleanCellValue()+" ");
			break;
	}
	
}
	System.out.println("");}
	
	}

}
