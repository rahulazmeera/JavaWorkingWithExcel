package excelreader;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class writingininexcelfile {

	/**
	 * @param args
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws EncryptedDocumentException 
	 */
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		 FileInputStream fiu=new FileInputStream("//home//rahul//Desktop//sheeet1.xlsx");
		 Workbook wb=WorkbookFactory.create(fiu);
		 Sheet s=wb.getSheet("Sheet1");
		 Row r=s.getRow(0);
		 Cell c=r.getCell(0);
		 c.setCellType(c.CELL_TYPE_NUMERIC);
		 c.setCellValue(140);
		 FileOutputStream fis=new FileOutputStream("//home//rahul//Desktop//sheeet1.xlsx");
		 wb.write(fis);
		 fis.close();
		 System.out.println("sucess");
		 
		
		
		
	
	}

}
