package SIM;

	import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
	import org.apache.poi.xssf.usermodel.XSSFWorkbook;

	import java.io.FileInputStream;
	import java.io.FileOutputStream;
	import java.io.IOException;

	public class datatransfer {

	    public static void main(String[] args) throws IOException {

	        // File paths
	        String filePath = ".\\Data\\documentSearch_rbharatm(3).xls";
	        String existingFilePath = ".\\Data\\Macros to send mail.xlsm";

	        // Input file
	        FileInputStream file = new FileInputStream(filePath);
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	        XSSFSheet sheet = workbook.getSheetAt(0);

	        // Existing file
	        FileInputStream existingFile = new FileInputStream(existingFilePath);
	        HSSFWorkbook existingWorkbook = new HSSFWorkbook(existingFile);
	        HSSFSheet existingSheet = existingWorkbook.getSheetAt(0);

	        // Transfer data
	        int existingRowNum = existingSheet.getLastRowNum();
	        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
	            org.apache.poi.xssf.usermodel.XSSFRow row = sheet.getRow(i);
	            org.apache.poi.hssf.usermodel.HSSFRow existingRow = existingSheet.createRow(existingRowNum + i + 1);
	            for (int j = 0; j < row.getLastCellNum(); j++) {
	                org.apache.poi.xssf.usermodel.XSSFCell cell = row.getCell(j);
	                org.apache.poi.hssf.usermodel.HSSFCell existingCell = existingRow.createCell(j);
	                existingCell.setCellValue(cell.getStringCellValue());
	            }
	        }

	        // Write to existing file
	        FileOutputStream outFile = new FileOutputStream(existingFilePath);
	        existingWorkbook.write(outFile);
	        outFile.close();

	        // Close input files
	        file.close();
	        existingFile.close();

	        System.out.println("Data transferred successfully!");
	    }
	}
