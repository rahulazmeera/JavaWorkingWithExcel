package excelreader;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Excelsheets {

	/**
	 * @param args
	 */
	 static int RowCount=0;
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		 FileInputStream fis=new FileInputStream("//home//rahul//Desktop//sheeet1.xlsx");
         Workbook wb=WorkbookFactory.create(fis); 
         org.apache.poi.ss.usermodel.Sheet s= wb.getSheet("Sheet1");
         
         RowCount =s.getLastRowNum();
         System.out.println(RowCount);
      
		
	}
	
}
