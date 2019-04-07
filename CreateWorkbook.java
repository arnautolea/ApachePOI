package mvn.ApachePOI;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateWorkbook {

		   public static void main(String[] args) {
		      
		      XSSFWorkbook wb = new XSSFWorkbook(); 
		      XSSFCreationHelper createHelper = wb.getCreationHelper();
		     
		      XSSFSheet sheet = wb.createSheet("Development, depId 1");
	          XSSFSheet sheet1 = wb.createSheet("Accounting, depId 2");
		     
		      XSSFRow row = sheet.createRow(0);
		      createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Emp ID"));
		      createCell(wb, row, 1, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Lastname"));
		      createCell(wb, row, 2, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Firstname"));
		      createCell(wb, row, 3, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Birthdate"));
		      createCell(wb, row, 4, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Manager ID"));
		      createCell(wb, row, 5, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Skills"));
		      makeBold(wb, sheet.getRow(0));
		      
		      XSSFRow row1 = sheet.createRow(1);
		      
		      row1.createCell(0).setCellValue(createHelper.createRichTextString("001"));
		      row1.createCell(1).setCellValue(createHelper.createRichTextString("Washington"));
		      row1.createCell(2).setCellValue(createHelper.createRichTextString("George"));
		      row1.createCell(3).setCellValue(createHelper.createRichTextString("February 22, 1732"));
		      row1.createCell(4).setCellValue(createHelper.createRichTextString("0"));
		      Cell cell   = row1.createCell(5);  
	            cell.setCellValue("Powers of persuasion \nAbility to unify \nEmpowering others");  
	            CellStyle cs = wb.createCellStyle();  
	            cs.setWrapText(true);  
	            cell.setCellStyle(cs);  
	            row1.setHeightInPoints((3*sheet.getDefaultRowHeightInPoints()));  
	            sheet.autoSizeColumn(3); 
	            
	         XSSFRow row2 = sheet.createRow(2);
			      
			  row2.createCell(0).setCellValue(createHelper.createRichTextString("002"));
			  row2.createCell(1).setCellValue(createHelper.createRichTextString("Adams"));
			  row2.createCell(2).setCellValue(createHelper.createRichTextString("John"));
			  row2.createCell(3).setCellValue(createHelper.createRichTextString("October 30, 1735"));
			  row2.createCell(4).setCellValue(createHelper.createRichTextString("001"));
			  Cell cell1   = row2.createCell(5);  
		        cell1.setCellValue("A great communicator \nSuccessful lawyer");  
		        CellStyle cs1 = wb.createCellStyle();  
		        cs1.setWrapText(true);  
		        cell1.setCellStyle(cs1);  
		        row2.setHeightInPoints((2*sheet.getDefaultRowHeightInPoints()));  
		        sheet.autoSizeColumn(2); 
		      
		        XSSFRow row3 = sheet1.createRow(0);
			      createCell(wb, row3, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Emp ID"));
			      createCell(wb, row3, 1, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Lastname"));
			      createCell(wb, row3, 2, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Firstname"));
			      createCell(wb, row3, 3, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Birthdate"));
			      createCell(wb, row3, 4, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Manager ID"));
			      createCell(wb, row3, 5, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM, ("Skills"));
			      makeBold(wb, sheet1.getRow(0));
			      
			      XSSFRow row4 = sheet1.createRow(1);
			      
			      row4.createCell(0).setCellValue(createHelper.createRichTextString("003"));
			      row4.createCell(1).setCellValue(createHelper.createRichTextString("Jefferson"));
			      row4.createCell(2).setCellValue(createHelper.createRichTextString("Thomas"));
			      row4.createCell(3).setCellValue(createHelper.createRichTextString("April 13, 1743"));
			      row4.createCell(4).setCellValue(createHelper.createRichTextString("001"));
			      Cell cell4   = row4.createCell(5);  
		            cell4.setCellValue
		            ("Sense of Justice \nGreat understanding of history and politics \nOpen mind for learning \nFollower of truth and reason");  
		            CellStyle cs4 = wb.createCellStyle();  
		            cs4.setWrapText(true);  
		            cell4.setCellStyle(cs4);  
		            row4.setHeightInPoints((4*sheet1.getDefaultRowHeightInPoints()));  
		            sheet1.autoSizeColumn(4); 
		            
		         XSSFRow row5 = sheet1.createRow(2);
				      
				  row5.createCell(0).setCellValue(createHelper.createRichTextString("004"));
				  row5.createCell(1).setCellValue(createHelper.createRichTextString("Madison"));
				  row5.createCell(2).setCellValue(createHelper.createRichTextString("James"));
				  row5.createCell(3).setCellValue(createHelper.createRichTextString("March 16, 1751"));
				  row5.createCell(4).setCellValue(createHelper.createRichTextString("003"));
				  Cell cell5   = row5.createCell(5);  
			        cell5.setCellValue("Knowledge of constitutionalism \nCritical thinking");  
			        CellStyle cs5 = wb.createCellStyle();  
			        cs5.setWrapText(true);  
			        cell5.setCellStyle(cs1);  
			        row5.setHeightInPoints((2*sheet1.getDefaultRowHeightInPoints()));  
			        sheet1.autoSizeColumn(2); 		      
		      
		      
		      try (OutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
		          wb.write(fileOut);
		         wb.close();
		         System.out.println("File changed");
		      
		     	} catch (FileNotFoundException e) {
					e.printStackTrace();
		     	} catch (IOException e1) {
					e1.printStackTrace();
				}
		   }
		   private static void createCell(Workbook wb, Row row, int column, HorizontalAlignment halign, VerticalAlignment valign, String value) {
		        Cell cell = row.createCell(column);
		        cell.setCellValue(value);
		        CellStyle cellStyle = wb.createCellStyle();
		        cellStyle.setAlignment(halign);
		        cellStyle.setVerticalAlignment(valign);
		        cell.setCellStyle(cellStyle);
		    }
		   public static void makeBold(Workbook workbook, Row row)
		   {
		       for (int rowIndex = 0; rowIndex < row.getLastCellNum(); rowIndex++)
		       {
		           Cell cell = row.getCell(rowIndex);
		           CellStyle cellStyle = cell.getCellStyle();
		           XSSFFont font = (XSSFFont) workbook.createFont();
		           font.setBold(true);
		           cellStyle.setFont(font);
		           cell.setCellStyle(cellStyle);
		       }
		   }
}
		    