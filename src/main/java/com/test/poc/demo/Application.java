package com.test.poc.demo;



import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.test.poc.demo.utils.CreateFile;


public class Application {

	public static void main(String[] args) throws IOException {
		Application obj=new Application();
		obj.createExcel();
		
	}
	
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
			
			ironManAnchor.setCol1(2); 
			ironManAnchor.setCol2(3); 
			ironManAnchor.setRow1(i); 
			ironManAnchor.setRow2(i+1); 
			
			drawing.createPicture(ironManAnchor, inputImagePictureID);
			
			
		}
		for (int j = 0; j < 4; j++) {
		    sheet.autoSizeColumn(j);
		}
		
		try (FileOutputStream saveExcel = new FileOutputStream("target/testDinesh.xlsx")) {
		    workbook.write(saveExcel);
		    
		   
		}
		/*
		 * ClassLoader classLoader = getClass().getClassLoader(); String path =
		 * classLoader.getResource("testDinesh.xlsx").getPath(); zipFile(path);
		 */
		zipFile("target/testDinesh.xlsx");
		 System.out.println("target/testDinesh.xlsx");
	}
	
	private static void zipFile(String filePath) {
        try {
            File file = new File(filePath);
            //String zipFileName = file.getName().concat(".zip");
            String zipFileName = "testDinesh.zip";
 
            FileOutputStream fos = new FileOutputStream(zipFileName);
            ZipOutputStream zos = new ZipOutputStream(fos);
 
            zos.putNextEntry(new ZipEntry(file.getName()));
 
            byte[] bytes = Files.readAllBytes(Paths.get(filePath));
            zos.write(bytes, 0, bytes.length);
            zos.closeEntry();
            zos.close();
 
        } catch (FileNotFoundException ex) {
            System.err.format("The file %s does not exist", filePath);
        } catch (IOException ex) {
            System.err.println("I/O error: " + ex);
        }
    }

}
