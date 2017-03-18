package excelLib;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class exel {

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		FileInputStream fis=new FileInputStream("D:\\sandip\\excelLib\\files\\test.xlsx");
		Workbook wb=WorkbookFactory.create(fis);
		
Sheet sh=wb.getSheet("Sheet1");
Row row=sh.getRow(0);
Cell c1=row.getCell(0);
System.out.println(c1.getStringCellValue());
int count=sh.getLastRowNum();
System.out.println(count);
System.out.println(sh.getFirstRowNum());
System.out.println(row.getFirstCellNum());
System.out.println(row.getLastCellNum());


Cell cell1=row.createCell(3);
cell1.setCellType(CellType.STRING);
FileOutputStream fos=new FileOutputStream("D:\\sandip\\excelLib\\files\\test.xlsx");
cell1.setCellValue("pass");
wb.write(fos);

	}

}
