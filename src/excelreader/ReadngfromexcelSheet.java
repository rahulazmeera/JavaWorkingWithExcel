package excelreader;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadngfromexcelSheet {

	/**
	 * @param args
	 */
	static int RowCount=0;
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		FileInputStream fil=new FileInputStream("//home//rahul//Desktop//sheeet1.xlsx");
		Workbook wb=WorkbookFactory.create(fil);
		org.apache.poi.ss.usermodel.Sheet s=wb.getSheet("Sheet1");
		Row row=s.getRow(2);
		Cell cl=row.getCell(0);
		System.out.println(cl);
		 String Cellvalue=cl.getStringCellValue();
		 System.out.print(Cellvalue);
		
		
		
		
		
		
		
		
		

	}

}
