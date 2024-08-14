//package SIM;
//
//import java.awt.AWTException;
//import java.awt.Robot;
//import java.awt.Toolkit;
//import java.awt.datatransfer.StringSelection;
//import java.awt.event.KeyEvent;
//import java.io.File;
//import java.io.FileInputStream;
////import java.io.FileNotFoundException;
//import java.io.IOException;
//import java.text.SimpleDateFormat;
//import java.time.Duration;
//import java.time.Instant;
//import java.time.LocalDate;
//import java.time.LocalDateTime;
//import java.time.ZoneId;
//import java.time.ZoneOffset;
//import java.time.format.DateTimeFormatter;
//import java.util.Date;
//
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellType;
//import org.apache.commons.io.FileUtils;
//import org.apache.poi.hslf.record.Document;
//import org.apache.poi.hssf.usermodel.*;
//import org.apache.poi.HSSF.usermodel.*;
//import org.jsoup.Jsoup;
//import org.openqa.selenium.By;
//import org.openqa.selenium.InvalidArgumentException;
//import org.openqa.selenium.JavascriptExecutor;
//import org.openqa.selenium.Keys;
//import org.openqa.selenium.WebDriver;
//import org.openqa.selenium.WebElement;
//import org.openqa.selenium.firefox.FirefoxDriver;
//import org.openqa.selenium.support.ui.ExpectedConditions;
//import org.openqa.selenium.support.ui.WebDriverWait;
//public class slamiss {
//
//	public static void main(String[] args) throws InterruptedException, IOException, AWTException {
//		// TODO Auto-generated method stub
////		
//		WebDriver driver=new FirefoxDriver();
//		driver.manage().window().maximize();
////		driver.get("https://issues.amazon.com/issues/search?q=status%3A(Open)+containingFolder%3A(2857712e-96fa-4b2c-8a01-0328fe175eb0)+-createDate%3A(%5BNOW-2DAYS..NOW%5D)&sort=lastUpdatedDate+desc&selectedDocument=148be64f-ac5d-41fd-a9ee-366e1186a33e");	
////		driver.findElement(By.id("user_name_field")).sendKeys("rbharatm");
////		driver.findElement(By.id("user_name_btn")).click();
////		Thread.sleep(3000);
////		driver.findElement(By.id("password_field")).sendKeys("Bh@rath2008");
////		driver.findElement(By.id("password_btn")).click();
////		//driver.findElement(By.id("user_name")).sendKeys("rbharatm");
////		//driver.findElement(By.id("password")).sendKeys("Bh@rath2008");
////		//driver.findElement(By.id("verify_btn")).click();
////		Thread.sleep(40000);
////		driver.findElement(By.xpath("//button[@class='btn btn-small dropdown-toggle']")).click();
////		driver.findElement(By.xpath("//a[@class='export-search-results']")).click();
////		//driver.findElement(By.id("add-all-export-columns")).click();
////		//driver.findElement(By.id("xls")).click();
////		driver.findElement(By.id("submit-custom-export-job")).click();
////		Thread.sleep(30000);
////		//driver.switchTo().frame(0);
////		driver.findElement(By.xpath("//a[@data-link='html{>fileName} href{:~getAttachmentURL(~jobId, id, stack)}']")).click();
////	
////		driver.findElement(By.cssSelector("body")).sendKeys(Keys.CONTROL + "t");
//		driver.get("https://ballard.amazon.com/owa/#path=/mail");
//		Thread.sleep(5000);
//		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
//		driver.findElement(By.xpath("//*[@title='Write a new message (N)']")).click();
//		Thread.sleep(30000);
//		driver.findElements(By.xpath("//*[@aria-label='To']")).get(0).sendKeys("rbharatm@amazon.com");
//		driver.findElements(By.xpath("//*[@placeholder='Add a subject']")).get(0).sendKeys("SLA MISS");
//		
//		//Attaching the latest downloaded file
//		File downloadDir = new File("C:\\Users\\rbharatm\\Downloads");
//		File[] files = downloadDir.listFiles();
//		File latestFile = null;
//		long latestTimestamp = 0;
//
//		for (File file : files) {
//		    if (file.lastModified() > latestTimestamp) {
//		        latestFile = file;
//		        latestTimestamp = file.lastModified();
//		    }
//		}
//
//		String downloadedFilePath = latestFile.getAbsolutePath();
//		
//		//Message body 
//		FileInputStream inputstream=new FileInputStream(downloadedFilePath);
//		
////		HSSFWorkbook workbook=new HSSFWorkbook(inputstream);
////		HSSFSheet sheet=workbook.getSheetAt(0);
//		
//		HSSFWorkbook workbook = new HSSFWorkbook(inputstream);
//		HSSFSheet sheet = workbook.getSheetAt(0);
//
//		int rows=sheet.getLastRowNum();
//		int cols=sheet.getRow(1).getLastCellNum();
//		
//		for(int r=0; r<=rows; r++) {
//		    HSSFRow row = sheet.getRow(r);
//		    if(row == null) {
//		        break; // skip empty rows
//		    }
//		    for(int c=0; c<row.getLastCellNum(); c++) {
//		        HSSFCell cell = row.getCell(c);
//		        if(cell == null) {
//		            System.out.print(" | "); // print separator for empty cells
//		            continue;
//		        }
//		        switch(cell.getCellType()) {
//		            case STRING:
//		                System.out.print(cell.getStringCellValue());
//		                break;
//		            case NUMERIC:
//		                System.out.print(cell.getNumericCellValue());
//		                break;
//		            case BOOLEAN:
//		                System.out.print(cell.getBooleanCellValue());
//		                break;
//		        }
//		        System.out.print(" | ");
//		    }
//		    System.out.println();
//		}
//		
////	     // Sort the rows based on the date column
////        for (int i = 1; i < rows; i++) { // Start from the second row (index 1)
////            for (int j = i + 1; j <= rows; j++) {
////                Row row1 = sheet.getRow(i);
////                Row row2 = sheet.getRow(j);
////                if (row1 != null && row2 != null) {
////                    Cell cell1 = row1.getCell(1); // assuming the date column is the second column
////                    Cell cell2 = row2.getCell(1);
////                    if (cell1 != null && cell2 != null) {
////                        double date1;
////                        double date2;
////                        if (cell1.getCellType() == CellType.NUMERIC) {
////                            date1 = cell1.getNumericCellValue();
////                        } else {
////                            date1 = Double.parseDouble(cell1.getStringCellValue());
////                        }
////                        if (cell2.getCellType() == CellType.NUMERIC) {
////                            date2 = cell2.getNumericCellValue();
////                        } else {
////                            date2 = Double.parseDouble(cell2.getStringCellValue());
////                        }
////                        if (date1 > date2) {
////                            // swap the rows
////                            sheet.removeRow(row1);
////                            sheet.shiftRows(i + 1, rows, -1);
////                            HSSFRow newRow = sheet.createRow(i);
////                            for (int k = 0; k < row2.getLastCellNum(); k++) {
////                                Cell cell = row2.getCell(k);
////                                if (cell != null) {
////                                    Cell newCell = newRow.createCell(k);
////                                    if (cell.getCellType() == CellType.FORMULA) {
////                                        newCell.setCellFormula(cell.getCellFormula());
////                                    } else if (cell.getCellType() == CellType.NUMERIC) {
////                                        newCell.setCellValue(cell.getNumericCellValue());
////                                    } else {
////                                        newCell.setCellValue(cell.getStringCellValue());
////                                    }
////                                }
////                            }
////                            sheet.removeRow(row2);
////                            sheet.shiftRows(j + 1, rows, -1);
////                        }
////                    }
////                }
////            }
////        }
//		
//		StringBuilder tableHtml = new StringBuilder("<p>Hi All,"
//				+ "<br>"
//				+ "Please find the data below for the missed audit:</p>"
//				+ "<br>"
//				+ "The date highlighted with yellow is issues more than 5 days from today."
//				+ "<br>"
//				+ "<br>" + "<table border='1'>");
//		
//
//		LocalDate today = LocalDate.now();
//		LocalDate fiveDaysFromNow = today.plusDays(5);
//
//		for (int i = 0; i <= rows; i++) {
//		    Row row = sheet.getRow(i);
//		    if (row != null) {
//		        tableHtml.append("<tr>");
//		        for (int j = 0; j < row.getLastCellNum(); j++) {
//		            Cell cell = row.getCell(j);
//		            if (cell != null) {
//		                if (i == 0) { // first row
//		                    tableHtml.append("<td style='font-weight: bold; text-align: center; border: 2px solid black; padding: 5px;'>");
//		                } else {
//		                    tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>");
//		                }
//		                switch (cell.getCellType()) {
//		                    case STRING:
//		                    	if (i == 6 && row.getCell(j).getStringCellValue().equals(sheet.getRow(4).getCell(j).getStringCellValue())) {
//		                            tableHtml.append("<span style='background-color: yellow; color: red;'>").append(cell.getStringCellValue()).append("</span>");
//		                        } else {
//		                            tableHtml.append(cell.getStringCellValue());
//		                        }
//		                        break;
//		                    case NUMERIC:
//		                    	double excelDate = cell.getNumericCellValue();
//		                    	LocalDateTime dateTime = LocalDateTime.ofInstant(Instant.ofEpochMilli((long) (excelDate - 25569) * 86400000), ZoneId.systemDefault());
//		                    	LocalDate date = dateTime.toLocalDate();
//		                    	LocalDate fiveDaysAgo = LocalDate.now().minusDays(5);
//		                    	if (date.isBefore(fiveDaysAgo)) {
//		                    	    tableHtml.append("<span style='background-color: yellow; color: red;'>").append(dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"))).append("</span>");
//		                    	} else {
//		                    	    tableHtml.append(dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
//		                    	}
////		                    	String formattedDateTime = dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
////		                    	tableHtml.append(formattedDateTime);
//		                    	break;
//
//		                    case BOOLEAN:
//		                        tableHtml.append(cell.getBooleanCellValue());
//		                        break;
//		                }
//		                tableHtml.append("</td>");
//		            } else {
//		                if (i == 0) { // first row
//		                    tableHtml.append("<td style='font-weight: bold; text-align: center; border: 2px solid black; padding: 5px;'>&nbsp;</td>");
//		                } else {
//		                    tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>&nbsp;</td>");
//		                }
//		            }
//		        }
//		        tableHtml.append("</tr>");
//		    }
//		}
//		tableHtml.append("</table>");
//
//
//
//
////		for (int i = 0; i <= rows; i++) {
////		    Row row = sheet.getRow(i);
////		    if (row != null) {
////		        tableHtml.append("<tr>");
////		        for (int j = 0; j < row.getLastCellNum(); j++) {
////		            Cell cell = row.getCell(j);
////		            if (cell != null) {
////		                if (i == 0) { // first row
////		                    tableHtml.append("<td style='font-weight: bold; text-align: center; border: 1px solid black;'>");
////		                } else {
////		                    tableHtml.append("<td style='text-align: center; border: 1px solid black;'>");
////		                }
////		                switch (cell.getCellType()) {
////		                    case STRING:
////		                        tableHtml.append(cell.getStringCellValue()).append("</td>");
////		                        break;
////		                    case NUMERIC:
////		                        double excelDate = cell.getNumericCellValue();
////		                        LocalDateTime dateTime = LocalDateTime.ofEpochSecond((long) (excelDate - 25569) * 24 * 60 * 60, 0, ZoneOffset.UTC);
////		                        tableHtml.append(dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"))).append("</td>");
////    							break;
////		                    case BOOLEAN:
////		                        tableHtml.append(cell.getBooleanCellValue()).append("</td>");
////		                        break;
////		                 
////		                }
////		            } else {
////		                if (i == 0) { // first row
////		                    tableHtml.append("<td style='font-weight: bold; text-align: center; border: 1px solid black;'></td>");
////		                } else {
////		                    tableHtml.append("<td style='text-align: center; border: 1px solid black;'></td>");
////		                }
////		            }
////		        }
////		        tableHtml.append("</tr>");
////		    }
////		}
//
//
//		tableHtml.append("</table>"
//				+ "<br>"
//				+ "<br>"
//				+ "<br>" + "<p>Thanks\n"
//				+ "<br>"
//				+ "Bharath M.</p>");
//		
////		System.out.println(tableHtml.toString());
//
////		org.jsoup.nodes.Document jsoupDoc = Jsoup.parse(tableHtml.toString());
//
//////		Document htmldoc = Jsoup.parse(jsoupDoc);
////		System.out.println(jsoupDoc.toString());
//
//
//
////		String plainText = Jsoup.parse(tableHtml.toString()).text();
////		driver.findElements(By.xpath("//*[@aria-label='Message body']")).get(0).sendKeys(plainText);
//
//		
////		driver.findElements(By.xpath("//*[@aria-label='Message body']")).get(0).sendKeys("Hi Dhivya,"
////				+ "\n"
////				+ "I have completed the script, please find the attachment for the sla miss as on 24/07/2024.\n"
////				+ "\n"
////				+ "Thanks \n"
////				+ "Bharath M. ");
//		
//		
////		driver.findElements(By.xpath("//*[@aria-label='Insert']")).get(0).click();
////		driver.findElements(By.xpath("//*[@title='Attach']")).get(0).click();
//		Thread.sleep(5000);
////		WebElement fileInput = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//input[@type='file']")));
////		try {
////		    fileInput.sendKeys("C:\\Users\\rbharatm\\Downloads\\documentSearch_rbharatm(4).txt");
////		} catch (InvalidArgumentException e) {
////		    System.out.println("File not found: " + e.getMessage());
////		}
////
////
//		
//		//option 2 
//		
//		
////		StringBuilder tableHtml = new StringBuilder("<p>Hi All," + "\n" + "Please find the data below for the missed audit:</p>"
////				+ "<p>\n<p>" + "<table border='1'>");
////
////		for (int r = 0; r <= rows; r++) {
////		    HSSFRow row = sheet.getRow(r);
////		    if (row == null) {
////		        break; // skip empty rows
////		    }
////		    tableHtml.append("<tr>");
////		    for (int c = 0; c < row.getLastCellNum(); c++) {
////		        HSSFCell cell = row.getCell(c);
////		        if (cell == null) {
////		            tableHtml.append("<td></td>"); // print separator for empty cells
////		            continue;
////		        }
////		        switch (cell.getCellType()) {
////		            case STRING:
////		                tableHtml.append("<td>").append(cell.getStringCellValue()).append("</td>");
////		                break;
////		            case NUMERIC:
////		            	tableHtml.append("<td>").append("'").append(cell.getNumericCellValue()).append("'</td>");
////		                break;
////
////		            case BOOLEAN:
////		                tableHtml.append("<td>").append(cell.getBooleanCellValue()).append("</td>");
////		                break;
////		            
////		        }
////		    }
////		    tableHtml.append("</tr>");
////		}
////		tableHtml.append("</table>");
////		
////		tableHtml.append("</table>" + "<p>\nThanks\n"
////				+ "\n"
////				+ "Bharath.M.</p>");
//
//		
//		String tableHtmlString = tableHtml.toString().replace("\n", "\\n");
//
//		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = `" + tableHtmlString + "`;", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
//
////		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = JSON.stringify('" + tableHtmlString + "');", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
//
////		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = '" + tableHtmlString + "';", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
//
////		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = JSON.stringify('" + tableHtml.toString() + "');", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
//		
////		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = '" + tableHtml.toString().replace("'", "\\'") + "';", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
//
////		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = JSON.stringify('" + tableHtml.toString() + "');", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
//
////		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = '" + tableHtml.toString() + "';", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
//
////		((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = '<html><body>This is a <b>test</b> email with HTML.</body></html>';", driver.findElement(By.xpath("//div[@aria-label='Message body']")));
//
//		// Attach the downloaded file\
//		
//		String newFileName = "SLA_Miss_Report_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xls";
//		File newFile = new File(downloadDir, newFileName);
//		latestFile.renameTo(newFile);
//		
//		String newdownloadedFilePath = newFile.getAbsolutePath();
//
//		WebElement fileInput = driver.findElement(By.xpath("//input[@type='file']"));
//		fileInput.sendKeys(newdownloadedFilePath);
//
////		
////		Robot robot = new Robot();
////		
////		String filePath = "C:\\Users\\rbharatm\\Downloads\\documentSearch_rbharatm.xls";
////		StringSelection selection = new StringSelection(filePath);
////		Toolkit.getDefaultToolkit().getSystemClipboard().setContents(selection, null);
////		robot.keyPress(KeyEvent.VK_CONTROL);
////        robot.keyPress(KeyEvent.VK_V);
////        robot.keyRelease(KeyEvent.VK_V);
////        robot.keyRelease(KeyEvent.VK_CONTROL);
////        robot.keyPress(KeyEvent.VK_ENTER);
////        robot.keyRelease(KeyEvent.VK_ENTER);
////       
//		Thread.sleep(10000);
////		driver.findElements(By.xpath("//*[@aria-label='Send']")).get(0).click();
//        Thread.sleep(10000);
////        driver.close();
//		
//
//		
////		FileInputStream inputstream=new FileInputStream(downloadedFilePath);
////		
////		HSSFWorkbook workbook=new HSSFWorkbook(inputstream);
////		HSSFSheet sheet=workbook.getSheetAt(0);
////		
////		int rows=sheet.getLastRowNum();
////		int cols=sheet.getRow(1).getLastCellNum();
////		
////		for(int r=0;r<=rows;r++)
////		{
////			HSSFRow row=sheet.getRow(r); 
////			
////			for(int c=0;c<=cols;c++) 
////			{
////				HSSFCell cell=row.getCell(c);
////				
////				switch(cell.getCellType())
////				{
////				case STRING: System.out.print(cell.getStringCellValue()); break;	
////				case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
////				case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
////				
////				}
////				System.out.println(" | ");
////			}
////			System.out.println();
////		}
//		
//		
//	}
//
//}
