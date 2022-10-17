package Excel.ReadwriteExcel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;


public class ExcelReading {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		String Path="C:\\Users\\Md Tarab\\eclipse-workspace\\SeleniumPractise\\DataDrivenFramework\\src\\test\\resources\\excel";

		
		try {
			FileInputStream fis= new FileInputStream(Path);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		XSSFWorkbook x= new XSSFWorkbook();
		XSSFSheet sh=x.getSheetAt(0);
		int rc=sh.getLastRowNum();
		int col=sh.getRow(1).getLastCellNum()  ;
		
		for(int i=0;i<=rc;i++)
		{
			XSSFRow r= sh.getRow(i);
			if(r==null)
				System.out.println("no row data found");
			for(int j=0;j<col;j++)
			{
				XSSFCell c=r.getCell(j);
				if(c==null)
				{
					System.out.println("no col data found");
				}
				else if(c.getCellType()==CellType.STRING)
				{
					System.out.println(c.getStringCellValue());
				}
				else if(c.getCellType()==CellType.NUMERIC)
				{
					System.out.println(c.getNumericCellValue());
				}
			}
		}
		///Through iterator
		
		Iterator itr=sh.iterator();
		while(itr.hasNext())
		{
			XSSFRow r=(XSSFRow)itr.next();
			Iterator cellitr=r.cellIterator();
			while(cellitr.hasNext())
			{
				XSSFCell c= (XSSFCell)cellitr.next();
				if(c==null)
				{
					System.out.println("no col data found");
				}
				else if(c.getCellType()==CellType.STRING)
				{
					System.out.println(c.getStringCellValue());
				}
				else if(c.getCellType()==CellType.NUMERIC)
				{
					System.out.println(c.getNumericCellValue());
				}
				
			}
		}
		
	}

}
