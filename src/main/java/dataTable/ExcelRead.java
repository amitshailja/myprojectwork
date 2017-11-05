package dataTable;

import java.io.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead{
		  XSSFWorkbook xworkbook;
		  XSSFSheet xsheet;
		  XSSFRow xrow;
	      XSSFCell xhcell,xdcell;
	      
	    private File folder =null;
	  	private File newFile=null;
	  	private String results;
	  	private int i;
	  	private ArrayList<String> al1 = new ArrayList();
	  	private ArrayList<String> h1 = new ArrayList();
	  	private Set<String> hs= new HashSet<String>();
	           
		public  void readExcelFile(String Filepath, int SheetIndex) {
				 try

			    {
				 FileInputStream in = new FileInputStream(Filepath);

			     File file = new File(Filepath);
			     ArrayList<String> header1  = new ArrayList();
			     ArrayList<String> data = new ArrayList();
			     LinkedHashMap<List, List> hashMap = new LinkedHashMap<List, List>();
			    

			    if(file.isFile() && file.exists()){

			        xworkbook = new XSSFWorkbook(in);
			        xsheet=xworkbook.getSheetAt(SheetIndex);
			       
			        int totalRows = xsheet.getLastRowNum();
	                int hcell =0;
	                int dcell =1;
	                
			        for(int row =0; row<=totalRows;row++){
			            xrow=xsheet.getRow(row);
			            int totalCells=xrow.getLastCellNum();

			        
			                if(xrow != null)
			                {
			                         xhcell = xrow.getCell(hcell); 
			                         xdcell = xrow.getCell(dcell);
			                            if(xhcell!=null)
			                         {
			                    	        cellvalueAString(xhcell,header1);	
			                                 
	                                 }
			                            
			                            if(xdcell!=null)
				                         {
				                    	        cellvalueAString(xdcell,data);	
				                                
		                                 }   
			                         		               
			            }    
			        }
			       
			    }
			    hashMap.put( header1, data);
			    wirteExcel("E:\\Shalu_Project\\Data Table.xlsx","QuoteCreation", header1);
			    ArrayList<String> fname = readFilename("E:\\Shalu_Project\\XMLs");
			    writeDatainExitingExcel(fname,data , "E:\\Shalu_Project\\Data Table.xlsx");
			
			   
			    
		    in.close();
			    
			}
			catch (Exception ex){
			    ex.printStackTrace();
			}
			 
		}		
		private void cellvalueAString(Cell cell, ArrayList<String> arr1) throws IOException {
			String cellvalue = "";
			switch (cell.getCellType()) {

			case Cell.CELL_TYPE_STRING:
				cellvalue = cell.getStringCellValue();
				arr1.add(cellvalue);
				break;

			case Cell.CELL_TYPE_NUMERIC:

				
				cellvalue = Integer.toString((int) cell.getNumericCellValue());
				arr1.add(cellvalue);
				break;
			
			case Cell.CELL_TYPE_BLANK:

				cellvalue = " ";
				arr1.add(cellvalue);
				break;

			}
			
		}
		
//***********************************************************************************************************************//
		
		public void wirteExcel(String FilePath,String SheetName,  ArrayList<String> headers) throws Exception
		{
			int rowidx =0;
			int cellidx = 0;
			
			//FileInputStream in = new FileInputStream(FilePath);
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet xs = wb.createSheet(SheetName);
			XSSFRow row = xs.createRow(rowidx);
			
			XSSFFont font = wb.createFont();
			font.setBold(true);
			font.setFontHeight(10);
			font.setFontName("Arial");
			font.setColor(IndexedColors.BLUE_GREY.getIndex());
			xs.autoSizeColumn(100000);
			
			
			for(cellidx =0; cellidx<headers.size();cellidx++){
				
			   
				XSSFCell cell = row.createCell(cellidx);
				CellStyle style = wb.createCellStyle();
				
				//style.setAlignment(CellStyle.ALIGN_CENTER);
				style.setWrapText(true);
				//style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				style.setFont(font);
		        cell.setCellStyle(style);
		        cell.setCellValue(headers.get(cellidx));
		        xs.autoSizeColumn(cellidx);
			}
		//	 in.close();
			 FileOutputStream  fos =new FileOutputStream( new File(FilePath));
	            wb.write(fos);
	           fos.close();
	        }
				

		//****************************Read File Name********************************************************************************//
		
		public ArrayList<String> readFilename(String FolderName) throws Exception
		{
			
			folder = new File(FolderName);
			File[] filelist = folder.listFiles();
			System.out.println(filelist.length);
			
			
			for(i=0;i<filelist.length;i++)
			{
				al1.add(filelist[i].getName().substring(0, filelist[i].getName().lastIndexOf(".")));	
			}
			
			hs.addAll(al1);
			al1.clear();
			al1.addAll(hs);
			for(i=0; i<al1.size();i++){
			System.out.println(al1.get(i));
			}
			return al1;
		}
		
	
		
		//********************* Write data into existing excel sheet *************************************//
		
	    public static void writeDatainExitingExcel(List<String> fname ,List<String> data , String updateFilePath) throws Exception {

			
			FileInputStream in = new FileInputStream(updateFilePath);
			FileOutputStream fos =null;
			 XSSFWorkbook workBook = new XSSFWorkbook(in);
			 XSSFSheet xsheet = workBook.getSheetAt(0);
			 XSSFRow row;
	         XSSFCell cell;
	        
	         for(int rowno =1; rowno<=fname.size();rowno++)
	        {
	        	 row = xsheet.createRow(rowno);
	        	 for(int cellno = 0;cellno<data.size();cellno++){
	        	 
	        	 cell= row.createCell(cellno);
	        	 if(cellno==2)
	        	 {
	        		 cell.setCellValue(fname.get(rowno-1).toString());
	        		 xsheet.autoSizeColumn(rowno);
	        	 }
	        	 else
	        	 cell.setCellValue(data.get(cellno));
	        	 xsheet.autoSizeColumn(cellno);}
	       }
	         	
	         in.close();
	         fos =new FileOutputStream( new File(updateFilePath));
	            workBook.write(fos);
	           fos.close();
	        }


	}
