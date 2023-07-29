package webdriver;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
public static void main(String[] args) throws IOException {
	
	
	FileInputStream file=new FileInputStream("E:\\Selenium\\input1.xlsx");
	XSSFWorkbook      wb=new XSSFWorkbook(file);
	XSSFSheet    sheet=wb.getSheet("Guru");
	
	
	int rows=sheet.getLastRowNum()- sheet.getFirstRowNum();
	
	
	XSSFRow row=sheet.getRow(1);
	XSSFCell cell=row.getCell(0);
	System.out.println(cell);
	
	for(int i=0;i<=rows;i++){
		Row r=sheet.getRow(i);
		for(int j=0;j<r.getLastCellNum();j++){
			
			System.out.println(sheet.getRow(i).getCell(j));
		}
	}
	
}
}
