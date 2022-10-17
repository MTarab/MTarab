package Excel.ReadwriteExcel;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriting {
	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook wb= new XSSFWorkbook();
		XSSFSheet sh=wb.createSheet("empoye");
		Object[][]arr= {{101,"red","black"},{102,"us","dhg"}};
		
		int row=arr.length;
		int col=arr[0].length;
		
		for(int i=0;i<row;i++)
		{
			XSSFRow r=sh.createRow(i);
			for(int j=0;j<col;j++)
			{
				XSSFCell c=r.createCell(j);
				Object value=arr[i][j];
				if(value instanceof String)
					c.setCellValue((String)value);
				else if(value instanceof Integer)
					c.setCellValue((Integer)value);
			}
		}
		
		String path="";
		FileOutputStream fo= new FileOutputStream(path);
		wb.write(fo);
		fo.close();
		
	}

}
