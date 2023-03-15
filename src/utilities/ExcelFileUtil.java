package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileUtil 
{
	XSSFWorkbook wb;
	//constructor for reading excel path
	public ExcelFileUtil(String excelpath)throws Throwable
	{
		FileInputStream fi=new FileInputStream(excelpath);
		wb=new XSSFWorkbook(fi);
	}
//count no of rows
	public int rowcount(String sheetName)
	{
		return wb.getSheet(sheetName).getLastRowNum();
	}
	//get cell type data

	public String getCellData(String sheetName,int row,int column)
	{
		String data="";
		if (wb.getSheet(sheetName).getRow(row).getCell(column).getCellType()==Cell.CELL_TYPE_NUMERIC) 
		{
			int celldata=(int)wb.getSheet(sheetName).getRow(row).getCell(column).getNumericCellValue();
			data=String.valueOf(celldata);	
		}
		else 
		{
			data=wb.getSheet(sheetName).getRow(row).getCell(column).getStringCellValue();
		}
		return data;
	}

		//method for set cell data
		public void  setCellData(String sheetName,int row,int column,String status,String writeexcel)throws Throwable
		{
		//get sheet from wb
			XSSFSheet ws=wb.getSheet(sheetName);
			//get row from sheet
			XSSFRow rowNum=ws.getRow(row);
			//create cell in row
			XSSFCell cell= rowNum.createCell(column);
			//write status
			cell.setCellValue(status);
			if (status.equalsIgnoreCase("pass")) 
			{
				XSSFCellStyle style=wb.createCellStyle();
				XSSFFont font=wb.createFont();
				font.setColor(IndexedColors.GREEN.getIndex());
				font.setBold(true);
				font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				style.setFont(font);
				rowNum.getCell(column).setCellStyle(style);
			}
			else if (status.equalsIgnoreCase("fail")) 
			{
				XSSFCellStyle style=wb.createCellStyle();
				XSSFFont font=wb.createFont();
				font.setColor(IndexedColors.RED.getIndex());
				font.setBold(true);
				font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				style.setFont(font);
				rowNum.getCell(column).setCellStyle(style);
			}
			else if (status.equalsIgnoreCase("blocked")) 
			{
				XSSFCellStyle style=wb.createCellStyle();
				XSSFFont font=wb.createFont();
				font.setColor(IndexedColors.BLUE.getIndex());
				font.setBold(true);
				font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
				style.setFont(font);
				rowNum.getCell(column).setCellStyle(style);
			}
			FileOutputStream fo=new FileOutputStream(writeexcel);
			wb.write(fo);	
		}
	
	public static void main(String[] args) throws Throwable
	{
		ExcelFileUtil xl=new ExcelFileUtil("D:/sample.xlsx");
		//count no of rows 
		int rc=xl.rowcount("empdata");
		System.out.println(rc);
		for (int i = 0; i <=rc; i++) 
		{
			String fname=xl.getCellData("empdata", i, 0);
			String mname=xl.getCellData("empdata", i, 1);
			String lname=xl.getCellData("empdata", i, 2);
			String eid=xl.getCellData("empdata", i, 3);
			System.out.println(fname+"  "+mname+"  "+lname+"  "+eid);
			xl.setCellData("EmpData", i, 4,"pass","D://Results.xlsx");
		}

	}

}