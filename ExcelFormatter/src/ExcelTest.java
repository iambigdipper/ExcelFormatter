import java.awt.Color;
import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTest {

	

	public ExcelTest() {
		// TODO Auto-generated constructor stub
	}
	
	//private static final String CSV_FILE_NAME = "C:\\Temp\\MyCSV.csv";
    //private static final String FILE_NAME = "C:\\Temp\\MyFirstExcel.xlsx";

    private static  String CSV_FILE_NAME = null;
    private static  String FILE_NAME = null;
    
    public static void main(String[] args) {

		/*
		 * XSSFWorkbook workbook = new XSSFWorkbook(); XSSFSheet sheet =
		 * workbook.createSheet("Datatypes in Java"); Object[][] datatypes = {
		 * {"Datatype", "Type", "Size(in bytes)"}, {"int", "Primitive", 2}, {"float",
		 * "Primitive", 4}, {"double", "Primitive", 8}, {"char", "Primitive", 1},
		 * {"String", "Non-Primitive", "Nooo fixxxed size"} };
		 * 
		 * int rowNum = 0; System.out.println("Creating excel");
		 * 
		 * 
		 * for (Object[] datatype : datatypes) {
		 * 
		 * XSSFCellStyle backgroundStyle = workbook.createCellStyle(); //XSSFColor
		 * myColor = new XSSFColor(Color.YELLOW);
		 * backgroundStyle.setFillBackgroundColor(new XSSFColor(Color.YELLOW, new
		 * DefaultIndexedColorMap()));
		 * backgroundStyle.setFillPattern(FillPatternType.LEAST_DOTS); Row row =
		 * sheet.createRow(rowNum++); //row.setRowStyle(backgroundStyle); int colNum =
		 * 0; for (Object field : datatype) {
		 * 
		 * 
		 * // backgroundStyle.setFillForegroundColor(new XSSFColor(Color.YELLOW, new
		 * DefaultIndexedColorMap()));//new XSSFColor(new java.awt.Color(128, 0, 128),
		 * new DefaultIndexedColorMap()) //backgroundStyle.set
		 * 
		 * Cell cell = row.createCell(colNum++);
		 * 
		 * if (field instanceof String) { cell.setCellValue((String) field); //
		 * cell.setCellStyle(arg0);
		 * 
		 * } else if (field instanceof Integer) { cell.setCellValue((Integer) field);
		 * 
		 * }
		 * 
		 * if(row.getRowNum()!=0 && row.getRowNum()%2==0)
		 * cell.setCellStyle(backgroundStyle);
		 * 
		 * //System.out.println(row.getRowStyle().getFillBackgroundColor());
		 * System.out.println(cell.getCellStyle().getFillBackgroundColor());
		 * //System.out.println(cell.getCellStyle().getFillForegroundColorColor());
		 * //cell.setCellStyle(backgroundStyle); } }
		 */
		/*
		 * try { FileOutputStream outputStream =convertcsvtoxls(); XSSFWorkbook workbook
		 * = new XSSFWorkbook(); workbook.write(outputStream); workbook.close();
		 * 
		 * convertcsvtoxls();
		 * 
		 * 
		 * } catch (FileNotFoundException e) { e.printStackTrace(); } catch (IOException
		 * e) { e.printStackTrace(); }
		 */
    	
    	CSV_FILE_NAME=args[0];
    	FILE_NAME=args[1];
        convertcsvtoxls();
        System.out.println("Done");
    }
    

    
    private static FileOutputStream convertcsvtoxls()
    {
    	 FileOutputStream fileOutputStream =null;
    	try {
            String csvFileAddress = CSV_FILE_NAME; //csv file address
            String xlsxFileAddress = FILE_NAME; //xlsx file address
            XSSFWorkbook workBook = new XSSFWorkbook();
       	   
       	    XSSFCellStyle  backgroundStyle = workBook.createCellStyle();
        	//XSSFColor myColor = new XSSFColor(Color.YELLOW);
            backgroundStyle.setFillBackgroundColor(new XSSFColor(Color.YELLOW, new DefaultIndexedColorMap()));
            backgroundStyle.setFillPattern(FillPatternType.LEAST_DOTS);
            
            XSSFSheet sheet = workBook.createSheet("sheet1");
            String currentLine=null;
            int RowNum=0;
            BufferedReader br = new BufferedReader(new FileReader(csvFileAddress));
            while ((currentLine = br.readLine()) != null) {
                String str[] = currentLine.split(",");
                RowNum++;
                XSSFRow currentRow=sheet.createRow(RowNum);
                for(int i=0;i<str.length;i++){
                    //currentRow.createCell(i).setCellValue(str[i]);
                    Cell cell = currentRow.createCell(i);
                    cell.setCellValue(str[i]);                  
                    if(currentRow.getRowNum()!=0 && currentRow.getRowNum()%2==0)
                    	  cell.setCellStyle(backgroundStyle);
                }
            }

            fileOutputStream =  new FileOutputStream(xlsxFileAddress);
            workBook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Done");
        } catch (Exception ex) {
            System.out.println(ex.getMessage()+"Exception in try");
        }
    	
    	return fileOutputStream;
    }
    

}
