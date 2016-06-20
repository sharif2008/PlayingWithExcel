

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.Queue;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.log4j.*;


public class ExcelParserUtility {
	@SuppressWarnings("resource")
	private static Logger logger =Logger.getLogger(ExcelParserUtility.class);
	public static void  processExcelFile(File file ) {
		int currentRowPosition=0;
		Workbook myWorkBook = null;
		try {
			// Creating Input Stream
			//FileInputStream myInput = new FileInputStream(file);
			
			// Create a workbook using the File System
			//XSSFWorkbook myWorkBook = new XSSFWorkbook(myInput);
			if (file.getName().toLowerCase().endsWith("xlsx")) {
				myWorkBook =WorkbookFactory.create(file);// new XSSFWorkbook(myInput);
		    } else if (file.getName().toLowerCase().endsWith("xls")) {
		    	myWorkBook =WorkbookFactory.create(file);// new HSSFWorkbook(myInput);
		    } else {
		        throw new IllegalArgumentException("The specified file is not Excel file");
		    }
			// Get the first sheet from workbook
			Sheet mySheet =  myWorkBook.getSheetAt(0);
			
			for(int i=1;i< myWorkBook.getNumberOfSheets();i++){
				myWorkBook.removeSheetAt(i);
			}
			
		

			/** We now need something to iterate through the cells. **/
			Iterator<Row> rowIter = mySheet.rowIterator();
			
			
			
			
			final DataFormatter dataFormatter = new DataFormatter();
			System.out.println("Staring Row ");
			
			while (rowIter.hasNext()) {
				Row myRow =  rowIter.next();
				Iterator<Cell> cellIter = myRow.cellIterator();
				
				while (cellIter.hasNext()) {
					Cell myCell = cellIter.next();
					
					
					//myCell.setCellType(1); //make all to String
					try{
						/*String cellStr =dataFormatter.formatCellValue(myCell).trim();
						System.out.println(cellStr);*/
						//Thread.sleep(1);   
						// TimeUnit.NANOSECONDS.sleep(100);
						System.out.println(myCell.getStringCellValue());
					}catch(Exception ex1){
						System.out.println(ex1);
						System.out.println("Cell column index: " + myCell.getColumnIndex());// get cell index
						System.out.println("Cell Type: " + myCell.getCellType());// get cell type
					}
					myCell=null;
					
				}
				myRow=null;
				//TimeUnit.NANOSECONDS.sleep(1);
				currentRowPosition++;
				
				System.out.println(currentRowPosition);
				
			}
			System.out.println("total rows: "+ currentRowPosition);
		}catch (Exception e) {
			e.printStackTrace();
		}finally {
			try{
				myWorkBook.close();
			}catch(Exception ex){
				System.out.println(ex);
			}
		}
	}
	
	
	public static Object getCellData(Cell cell) {
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				return cell.getRichStringCellValue().getString();
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					return cell.getDateCellValue();
				} else {
					return cell.getNumericCellValue();
				}
	
			default:
				return null;
		}
		

	}
	public static void main(String[] args) {
		 File xlsxFile = new File("test.xlsx");
		 ExcelParserUtility ex = new  ExcelParserUtility();
		 
		 ex.processExcelFile(xlsxFile);
	}
}
