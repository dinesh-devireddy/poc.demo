package com.test.poc.demo.utils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//import com.test.poc.demo.model.Customer;

public class CreateFile {
	
	
	
	//private List<Customer>customers;
	
	public void createExcel() throws IOException {
		
		XSSFSheet sheet;
		XSSFWorkbook workbook;
		//this.customers=customers;
		workbook = new XSSFWorkbook();
		sheet= workbook.createSheet("CustomersData");
		
		writeHeaderRow(sheet,workbook);
		
	}
	
	public void writeHeaderRow(XSSFSheet sheet, XSSFWorkbook workbook) throws IOException {
		Row row=sheet.createRow(0);
		
		Cell cell=row.createCell(0);
		cell.setCellValue("CustomerID");
		
		Cell cell2=row.createCell(1);
		cell2.setCellValue("CustomerName");
		
		Cell cell3=row.createCell(2);
		cell3.setCellValue("CustomerPhoto");
		
		
		writeDataRows(sheet,workbook);
		
		
	}
	
	public void writeDataRows(XSSFSheet sheet, XSSFWorkbook workbook) throws IOException {
		
		for(int i=1;i<=200;i++) {
			
			Row row=sheet.createRow(i);
			
			Cell cell=row.createCell(0);
			cell.setCellValue("id");
			
			Cell cell2=row.createCell(1);
			cell2.setCellValue("name");
			
			InputStream inputStream = CreateFile.class.getClassLoader()
				    .getResourceAsStream("ironman.png");
			
			byte[] inputImageBytes = IOUtils.toByteArray(inputStream);
			
			int inputImagePictureID = workbook.addPicture(inputImageBytes, Workbook.PICTURE_TYPE_PNG);
			
			XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
			
			XSSFClientAnchor ironManAnchor = new XSSFClientAnchor();
			
			ironManAnchor.setCol1(1); 
			ironManAnchor.setCol2(2); 
			ironManAnchor.setRow1(0); 
			ironManAnchor.setRow2(i); 
			
			drawing.createPicture(ironManAnchor, inputImagePictureID);
			
			
		}
		for (int j = 0; j < 3; j++) {
		    sheet.autoSizeColumn(j);
		}
		
		try (FileOutputStream saveExcel = new FileOutputStream("target/testDinesh.xlsx")) {
		    workbook.write(saveExcel);
		}
	}
}
