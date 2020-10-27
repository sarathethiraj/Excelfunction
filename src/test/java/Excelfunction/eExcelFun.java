package Excelfunction;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class eExcelFun {
	
	static String  value;
	static Workbook wb;
	static Sheet sh;
	static DataFormatter df ;
	
	public static void readxl() throws IOException{
		FileInputStream fis= new FileInputStream("E://samplefile.xlsx");
		 wb=WorkbookFactory.create(fis);
		 sh= wb.getSheet("Sheet1");
		int row=sh.getPhysicalNumberOfRows();
		int col=sh.getRow(0).getPhysicalNumberOfCells();
		System.out.println("Rownumber"+row);
		System.out.println("colnumber"+col);
		 df = new DataFormatter();
		 HashMap<String, String> mp = new HashMap<String, String>();
		for(int i=0; i<row; i++){
			for(int j=0;j<col;j++){
				
			mp.put(sh.getRow(0).getCell(j).getStringCellValue(), cellvalue(i,j));
		
				
		}
			System.out.println(mp);
		}
		
		
		
		
		
		
		
	}
		
		
		public static String cellvalue(int rownum, int colnum){
			
             Cell cel= sh.getRow(rownum).getCell(colnum);
            
             if(df.formatCellValue(cel).isEmpty()){
	            	value= "";
	            }
			
             else if(cel.getCellType()==CellType.NUMERIC){
				
				
				  value=df.formatCellValue(cel);
				
			
			
			}
			
			else if(cel.getCellType()==CellType.FORMULA){
				 double val=cel.getNumericCellValue();
				 value= String.valueOf(val);
			}
			
			
			else{
				
				
				 value=df.formatCellValue(cel);
			}
			
			//System.out.println(value );	
			return value;
			}
		
		
	
	

	public static void main(String[] args) throws IOException {
		
            readxl();
	}

}
