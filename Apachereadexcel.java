package guru;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Apachereadexcel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fn= new FileInputStream("C:\\Users\\vinishka pranav\\Desktop\\Book1.xlsx");
		XSSFWorkbook wrkbok= new XSSFWorkbook(fn);
		XSSFSheet sht=wrkbok.getSheet("Sheet1");
		
		int row=sht.getLastRowNum();
	    int cell=sht.getRow(1).getLastCellNum();
	    
	    for (int i=0;i<=row;i++) {
	    	
	    	XSSFRow rt = sht.getRow(i);
	    	
	    	for(int c=0;c<cell;c++)
	    	{
	    		XSSFCell ct= rt.getCell(c);
	    		switch(ct.getCellType()) {
	    		
	    		case STRING:
	    			System.out.print(ct.getStringCellValue()+" ");
	    			break;
	    		case NUMERIC:
	    			System.out.print(ct.getNumericCellValue()+" ");
	    			break;
	    		case BOOLEAN:
	    			System.out.print(ct.getBooleanCellValue()+" ");
	    			break;
	    		}
	    		}
	    	System.out.println(""); }

	}

}
