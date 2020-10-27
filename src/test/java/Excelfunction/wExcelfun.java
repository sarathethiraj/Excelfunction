package Excelfunction;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;

public class wExcelfun {
	static String  value;
	static Workbook wb;
	static Sheet sh;
	static DataFormatter df ;
	static CellAddress celladd;
	
	public static void value(String add, int val){
		
		 celladd= new CellAddress(add);
		sh.getRow(celladd.getRow()).getCell(celladd.getColumn()).setCellValue(val);
		}
public static void value(String add, String val){
		
		 celladd= new CellAddress(add);
	
		sh.getRow(celladd.getRow()).getCell(celladd.getColumn()).setCellValue(val);
		}
public static void value(String add, double val){
	
	 celladd= new CellAddress(add);
	sh.getRow(celladd.getRow()).getCell(celladd.getColumn()).setCellValue(val);
	}
	
	public static void rdfun() throws EncryptedDocumentException, IOException{
		FileInputStream fis= new FileInputStream("E://rate.xls");
		wb=WorkbookFactory.create(fis);
		
		sh=wb.getSheet("Sheet1");
		value("A3",250);
		value("A4",1.5);
		value("A1","Sarath");
		wb.setForceFormulaRecalculation(true);
		fis.close();
		FileOutputStream fos= new FileOutputStream("E://rate2.xls");
		wb.write(fos);
		System.out.println("done");
		fos.close();
	}
	

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		rdfun();
	}

}
