package SIM;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Audit_WBR {
	
	 public static void main(String[] args) throws InterruptedException, IOException, AWTException {
	        WebDriver driver = new FirefoxDriver();
	        driver.manage().window().maximize();
//			driver.get("https://issues.amazon.com/issues/search?q=containingFolder%3A(2a42d4ee-18c1-4cf8-8a1e-b1344ff622c8)+createDate%3A(%5BNOW-7DAYS..NOW%5D)&sort=lastUpdatedDate+desc&selectedDocument=4bcf3a6b-892f-43e6-a53b-5f19007c3ddb");	
//			//Enter username in sendkeys
//			driver.findElement(By.id("user_name_field")).sendKeys("rbharatm");
//			driver.findElement(By.id("user_name_btn")).click();
//			Thread.sleep(3000);
//			//Enter password in sendkeys
//			driver.findElement(By.id("password_field")).sendKeys("Bh@rath2008");
//			driver.findElement(By.id("password_btn")).click();
//			//driver.findElement(By.id("user_name")).sendKeys("");
//			//driver.findElement(By.id("password")).sendKeys("");
//			//driver.findElement(By.id("verify_btn")).click();
//			Thread.sleep(40000);
//			driver.findElement(By.xpath("//button[@class='btn btn-small dropdown-toggle']")).click();
//			driver.findElement(By.xpath("//a[@class='export-search-results']")).click();
//			driver.findElement(By.id("submit-custom-export-job")).click();
//			Thread.sleep(30000);
//			driver.findElement(By.xpath("//a[@data-link='html{>fileName} href{:~getAttachmentURL(~jobId, id, stack)}']")).click();
//		
//			driver.findElement(By.cssSelector("body")).sendKeys(Keys.CONTROL + "t");

	        driver.get("https://ballard.amazon.com/owa/#path=/mail");
	        Thread.sleep(5000);
	        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	        driver.findElement(By.xpath("//*[@title='Write a new message (N)']")).click();
	        Thread.sleep(30000);
	        //Enter the userid and groups ex:rbharatm@amazon.com
	        driver.findElements(By.xpath("//*[@aria-label='To']")).get(0).sendKeys("tdhivya@amazon.com ; rbharatm@amazon.com");
	        //Enter the subject to be entered in the mail.
	        driver.findElements(By.xpath("//*[@placeholder='Add a subject']")).get(0).sendKeys("Audit summary for Execution and Non-Execution 7 Days");

	        File downloadDir = new File("C:\\Users\\rbharatm\\Downloads");
	        File[] files = downloadDir.listFiles();
	        if (files == null) {
	            throw new IOException("No files found in the download directory.");
	        }
	        
	        File latestFile = null;
	        long latestTimestamp = 0;
	        for (File file : files) {
	            if (file.isFile() && file.lastModified() > latestTimestamp) {
	                latestFile = file;
	                latestTimestamp = file.lastModified();
	            }
	        }
	        
	        if (latestFile == null) {
	            throw new IOException("No valid files found in the download directory.");
	        }
	        
	        String downloadedFilePath = latestFile.getAbsolutePath();
	        FileInputStream inputstream = new FileInputStream(downloadedFilePath);
	        HSSFWorkbook workbook = new HSSFWorkbook(inputstream);
	        HSSFSheet sheet = workbook.getSheetAt(0);
	        
	        // Extract rows and sort them by the created date column (index 1)
	        List<Row> rowList = new ArrayList<>();
	        for (int r = 0; r <= sheet.getLastRowNum(); r++) { // Adjusted to include all rows
	            Row row = sheet.getRow(r);
	            if (row != null) {
	                rowList.add(row);
	            }
	        }
	        
	        // Sort the rows based on the created date column (index 1)
	        Collections.sort(rowList, new Comparator<Row>() {
	            @Override
	            public int compare(Row row1, Row row2) {
	                LocalDate date1 = getDateFromCell(row1.getCell(1));
	                LocalDate date2 = getDateFromCell(row2.getCell(1));
	                return date1.compareTo(date2);
	            }

	            private LocalDate getDateFromCell(Cell cell) {
	                if (cell == null) {
	                    return LocalDate.MIN;
	                }
	                try {
	                    if (cell.getCellType() == CellType.NUMERIC) {
	                        if (DateUtil.isCellDateFormatted(cell)) {
	                            return cell.getLocalDateTimeCellValue().toLocalDate();
	                        } else {
	                            double excelDate = cell.getNumericCellValue();
	                            return LocalDate.ofEpochDay((long) (excelDate - 25569));
	                        }
	                    } else if (cell.getCellType() == CellType.STRING) {
	                        String dateStr = cell.getStringCellValue();
	                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
	                        return LocalDate.parse(dateStr, formatter);
	                    }
	                } catch (Exception e) {
	                    System.out.println("Error parsing date from cell: " + e.getMessage());
	                }
	                return LocalDate.MIN;
	            }
	        });
	        //Enter the body of the mail. Use <br> to go to next line. 
	        StringBuilder tableHtml = new StringBuilder("<p>Hi All," + "<br><br>" + "Please find the data below for the audit summary:</p>" +  "<br><br>");

	        LocalDate today = LocalDate.now();
	        LocalDate fiveDaysAgo = today.minusDays(5);

	        // Summary section
	        tableHtml.append("<p style='font-size: 18px; font-weight: bold;'>Summary:</p><br>");
	     // Count of create date for last month
	        int AuditCount = 0;
	        YearMonth currentMonth = YearMonth.now();
	        YearMonth lastMonth = currentMonth.minusMonths(1);
	        for (Row row : rowList) {
	            if (row != null) {
	                Cell createDateCell = row.getCell(1);
	                if (createDateCell != null) {
	                    if (createDateCell.getCellType() == CellType.NUMERIC) {
	                        LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
	                        if (createDate.getYear() == lastMonth.getYear()  ) { //&& createDate.getMonthValue() == lastMonth.getMonthValue()
	                            AuditCount++;
	                        }
	                    }
	                }
	            }
	        }
	        //&& createDate.getMonthValue() == lastMonth.getMonthValue()
	        tableHtml.append("<p style='font-weight: bold;'>Total No.of Audit in the 7 days : " + AuditCount + "</p>");
	        tableHtml.append("</p><br>");
	        
	     // Types of Audit section
//	        tableHtml.append("<p style= 'font-weight: bold;'>Types of Audit Resolved:</p>");
//	        tableHtml.append("<table border='1' style='text-align: center;'>");
//	        tableHtml.append("<tr style='background-color: lightblue;'>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>SL.NO</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Audit Type</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Count</td>");
//	        tableHtml.append("</tr>");
//
//	        int executionAuditCount = 0;
//	        int nonExecutionAuditCount = 0;
//	        int bugAuditCount = 0;
//	        int slaMissCount = 0;
//	        int devicebrick = 0;
//	        int serialNumber = 1;
//
//	        for (Row row : rowList) {
//	            if (row != null) {
//	                Cell typeCell = row.getCell(3); // Assuming the type of audit is in the 3rd column
//	                if (typeCell != null) {
//	                    String type = typeCell.getStringCellValue();
//	                    if (type.contains("Execution Audit")) {
//	                        executionAuditCount++;
//	                    } else if (type.contains("Running Audit") || type.contains("Park Audit")) {
//	                        nonExecutionAuditCount++;
//	                    } else if (type.contains("Bug Audit")) {
//	                        bugAuditCount++;
//	                    } else if (type.contains("SLA")) {
//	                        slaMissCount++;
//	                    }else if (type.contains("Device Bricking")) {
//	                    	devicebrick++;
//	                    }
//	                }
//	            }
//	        }
//
//
//	        tableHtml.append("<tr>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber++).append("</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Execution Audit</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(executionAuditCount).append("</td>");
//	        tableHtml.append("</tr>");
//
//	        tableHtml.append("<tr>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber++).append("</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Non - Execution Audit</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nonExecutionAuditCount).append("</td>");
//	        tableHtml.append("</tr>");
//
//	        tableHtml.append("<tr>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber++).append("</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Bug Audit</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(bugAuditCount).append("</td>");
//	        tableHtml.append("</tr>");
//
//	        tableHtml.append("<tr>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber++).append("</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>SLA - Miss</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(slaMissCount).append("</td>");
//	        tableHtml.append("</tr>");
//	        
//	        tableHtml.append("<tr>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber++).append("</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Device Brick</td>");
//	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(devicebrick).append("</td>");
//	        tableHtml.append("</tr>");
//
//	        tableHtml.append("</table><br>");
	        
	        //Audit table
	        tableHtml.append("<p style= 'font-weight: bold;'>Types of Deviation:</p>");
	        tableHtml.append("<table border='1' style='text-align: center;'>");
	        tableHtml.append("<tr style='background-color: lightblue;'>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>SL.NO</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Audit Type</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>NC1</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>NC2</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>UserError</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Waivedoff</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Grand Total</td>");
	        tableHtml.append("</tr>");
	        
			int nc1ExecutionAuditCount = 0;
			int nc2ExecutionAuditCount = 0;
			int userErrorExecutionAuditCount = 0;
			int waiveoffExecutionAuditCount = 0;
			int totalExecutionAuditCount = 0;
			
			int nc1NonExecutionAuditCount = 0;
			int nc2NonExecutionAuditCount = 0;
			int userErrorNonExecutionAuditCount = 0;
			int waiveoffNonExecutionAuditCount = 0;
			int totalNonExecutionAuditCount = 0;
			
			int nc1BugAuditCount = 0;
			int nc2BugAuditCount = 0;
			int userErrorBugAuditCount = 0;
			int waiveoffBugAuditCount = 0;
			int totalBugAuditCount = 0;
			
			int nc1QueryCount = 0;
			int nc2QueryCount = 0;
			int userErrorQueryCount = 0;
			int waiveoffQueryCount = 0;
			int totalQueryCount = 0;
			
			int nc1DeviceBrickCount = 0;
			int nc2DeviceBrickCount = 0;
			int userErrorDeviceBrickCount = 0;
			int waiveoffDeviceBrickCount = 0;
			int totalDeviceBrickCount = 0;
			
			int serialNumber3 = 1;
		        for (Row row : rowList) {
		            if (row != null) {
		                Cell createDateCell = row.getCell(1);
		                if (createDateCell != null) {
		                    if (createDateCell.getCellType() == CellType.NUMERIC) {
		                        LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
		                        if (createDate.getYear() == lastMonth.getYear()  ){
			 { // Start from the second row (index 1)
			    if (row != null) {
			        Cell thirdColumnCell = row.getCell(3); // 4th column (index 3)
			        Cell fourthColumnCell = row.getCell(4); // 5th column (index 4)
			
			        if (thirdColumnCell != null && fourthColumnCell != null) {
			    String thirdColumnValue = thirdColumnCell.getStringCellValue();
			    String fourthColumnValue = fourthColumnCell.getStringCellValue();
			
			    // Execution Audit counts
			    if (thirdColumnValue.contains("Execution Audit")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1ExecutionAuditCount++;
			            System.out.println(nc1ExecutionAuditCount);
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2ExecutionAuditCount++;
			            System.out.println(nc2ExecutionAuditCount);
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorExecutionAuditCount++;
			            System.out.println(userErrorExecutionAuditCount);
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffExecutionAuditCount++;
			        }
			    }
			
			    // Non-Execution Audit counts
			    if (thirdColumnValue.contains("Running Audit") || thirdColumnValue.contains("Park Audit") || thirdColumnValue.contains("SLA") || thirdColumnValue.contains("Query ") || thirdColumnValue.contains("Bug Audit")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1NonExecutionAuditCount++;
			            System.out.println(nc1NonExecutionAuditCount);
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2NonExecutionAuditCount++;
			            System.out.println(userErrorNonExecutionAuditCount);
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorNonExecutionAuditCount++;
			            System.out.println(userErrorNonExecutionAuditCount);
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffNonExecutionAuditCount++;
			        }
			    }
			
			   
			
			    // Device Brick counts
			    if (thirdColumnValue.contains("Device Brick")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1DeviceBrickCount++;
			            System.out.println(nc1DeviceBrickCount);
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2DeviceBrickCount++;
			            System.out.println(nc2DeviceBrickCount);
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorDeviceBrickCount++;
			            System.out.println(userErrorDeviceBrickCount);
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffDeviceBrickCount++;
			        }
			    }
			}
			    }
			}
		                    }
		                }
		            }
		        }
		        }
			  
			int totalExecutionAuditRow = nc1ExecutionAuditCount + nc2ExecutionAuditCount + userErrorExecutionAuditCount + waiveoffExecutionAuditCount;
			int totalNonExecutionAuditRow = nc1NonExecutionAuditCount + nc2NonExecutionAuditCount + userErrorNonExecutionAuditCount + waiveoffNonExecutionAuditCount;
			int totalBugAuditRow = nc1BugAuditCount + nc2BugAuditCount + userErrorBugAuditCount + waiveoffBugAuditCount;
			int totalQueryRow = nc1QueryCount + nc2QueryCount + userErrorQueryCount + waiveoffQueryCount;
			int totalDeviceBrickRow = nc1DeviceBrickCount + nc2DeviceBrickCount + userErrorDeviceBrickCount + waiveoffDeviceBrickCount;
		
			
			int totalnc1 = nc1BugAuditCount + nc1DeviceBrickCount + nc1ExecutionAuditCount + nc1NonExecutionAuditCount + nc1QueryCount;
			int totalnc2 = nc2BugAuditCount + nc2DeviceBrickCount + nc2ExecutionAuditCount + nc2NonExecutionAuditCount + nc2QueryCount;
			int totaluserError = userErrorBugAuditCount + userErrorDeviceBrickCount + userErrorExecutionAuditCount + userErrorNonExecutionAuditCount + userErrorQueryCount;
			int totalwaiveoff = waiveoffExecutionAuditCount + waiveoffNonExecutionAuditCount + waiveoffBugAuditCount + waiveoffQueryCount + waiveoffDeviceBrickCount;
			int total = totalExecutionAuditRow + totalNonExecutionAuditRow + totalBugAuditRow + totalQueryRow + totalDeviceBrickRow ;
			
			// Execution Audit
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber3++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Execution Audit</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1ExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2ExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalExecutionAuditRow).append("</td>");
			tableHtml.append("</tr>");

			// Non-Execution Audit
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber3++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Non - Execution Audit</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1NonExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2NonExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorNonExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffNonExecutionAuditCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalNonExecutionAuditRow).append("</td>");
			tableHtml.append("</tr>");

			// Device Brick
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber3++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Device Brick</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1DeviceBrickCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2DeviceBrickCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorDeviceBrickCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffDeviceBrickCount).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalDeviceBrickRow).append("</td>");
			tableHtml.append("</tr>");
			
			//Total
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>Grand Total </td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalnc1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalnc2).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totaluserError).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalwaiveoff).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(total).append("</td>");
			tableHtml.append("</tr>");
			
			tableHtml.append("</table><br>");
			
			
			
			  //Audit table
	        tableHtml.append("<p style= 'font-weight: bold;'>Types of Deviation:</p>");
	        tableHtml.append("<table border='1' style='text-align: center;'>");
	        tableHtml.append("<tr style='background-color: lightblue;'>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>SL.NO</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Audit Type</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>NC1</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>NC2</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>UserError</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Waivedoff</td>");
	        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>Total</td>");
	        tableHtml.append("</tr>");
	        //Execution
			int nc1ExecutionAuditCount1 = 0;
			int nc2ExecutionAuditCount1 = 0;
			int userErrorExecutionAuditCount1 = 0;
			int waiveoffExecutionAuditCount1 = 0;
			int totalExecutionAuditCount1 = 0;
			//Non-Execution
			int nc1NonExecutionAuditCount1 = 0;
			int nc2NonExecutionAuditCount1 = 0;
			int userErrorNonExecutionAuditCount1 = 0;
			int waiveoffNonExecutionAuditCount1 = 0;
			int totalNonExecutionAuditCount1 = 0;
			//Park
			int nc1ParkCount1 = 0;
			int nc2ParkCount1 = 0;
			int userErrorParkCount1 = 0;
			int waiveoffParkCount1 = 0;
			int totalParkCount1 = 0;
			//Time sheet
			int nc1TimeSheetCount1 = 0;
			int nc2TimeSheetCount1 = 0;
			int userErrorTimeSheetCount1 = 0;
			int waiveoffTimeSheetCount1 = 0;
			int totalTimeSheetCount1 = 0;
			//Build
			int nc1BuildCount1 = 0;
			int nc2BuildCount1 = 0;
			int userErrorBuildCount1 = 0;
			int waiveoffBuildCount1 = 0;
			int totalBuildCount1 = 0;
			//Log
			int nc1LogCount1 = 0;
			int nc2LogCount1 = 0;
			int userErrorLogCount1 = 0;
			int waiveoffLogCount1 = 0;
			int totalLogCount1 = 0;
			//Retest
			int nc1RetestCount1 = 0;
			int nc2RetestCount1 = 0;
			int userErrorRetestCount1 = 0;
			int waiveoffRetestCount1 = 0;
			int totalRetestCount1 = 0;
			//Clarification query
			int nc1ClarificationQueryCount1 = 0;
			int nc2ClarificationQueryCount1 = 0;
			int userErrorClarificationQueryCount1 = 0;
			int waiveoffClarificationQueryCount1 = 0;
			int totalClarificationQueryCount1 = 0;
			//Running
			int nc1RunningCount1 = 0;
			int nc2RunningCount1 = 0;
			int userErrorRunningCount1 = 0;
			int waiveoffRunningCount1 = 0;
			int totalRunningCount1 = 0;
			//Bug
			int nc1BugAuditCount1 = 0;
			int nc2BugAuditCount1 = 0;
			int userErrorBugAuditCount1 = 0;
			int waiveoffBugAuditCount1 = 0;
			int totalBugAuditCount1 = 0;
			//Query
			int nc1QueryCount1 = 0;
			int nc2QueryCount1 = 0;
			int userErrorQueryCount1 = 0;
			int waiveoffQueryCount1 = 0;
			int totalQueryCount1 = 0;
			//Device brick
			int nc1DeviceBrickCount1 = 0;
			int nc2DeviceBrickCount1 = 0;
			int userErrorDeviceBrickCount1 = 0;
			int waiveoffDeviceBrickCount1 = 0;
			int totalDeviceBrickCount1 = 0;
			
			int serialNumber4 = 1;
			for (Row row : rowList) {
	            if (row != null) {
	                Cell createDateCell = row.getCell(1);
	                if (createDateCell != null) {
	                    if (createDateCell.getCellType() == CellType.NUMERIC) {
	                        LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
	                        if (createDate.getYear() == lastMonth.getYear()  ){ { // Start from the second row (index 1)
			    if (row != null) {
			        Cell thirdColumnCell = row.getCell(3); // 4th column (index 3)
			        Cell fourthColumnCell = row.getCell(4); // 5th column (index 4)
			
			        if (thirdColumnCell != null && fourthColumnCell != null) {
			    String thirdColumnValue = thirdColumnCell.getStringCellValue();
			    String fourthColumnValue = fourthColumnCell.getStringCellValue();
			
			    // Execution Audit counts
			    if (thirdColumnValue.contains("Execution Audit")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1ExecutionAuditCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2ExecutionAuditCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorExecutionAuditCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffExecutionAuditCount1++;
			        }
			    }
			
			    // Non-Execution Audit counts
			    if (thirdColumnValue.contains("SLA")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1NonExecutionAuditCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2NonExecutionAuditCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorNonExecutionAuditCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffNonExecutionAuditCount1++;
			        }
			    }
			    
				//Park
			    if (thirdColumnValue.contains("Park")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1ParkCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2ParkCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorParkCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffParkCount1++;
			        }
			    }
				//Time sheet
			    if (thirdColumnValue.contains("Time Sheet")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1TimeSheetCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2TimeSheetCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorTimeSheetCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffTimeSheetCount1++;
			        }
			    }
				//Build
			    if (thirdColumnValue.contains("Build")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1BuildCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2BuildCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorBuildCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffBuildCount1++;
			        }
			    }
			    
				//Log
			    if (thirdColumnValue.contains("Log")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1LogCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2LogCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorLogCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffLogCount1++;
			        }
			    }
				//Retest
			    if (thirdColumnValue.contains("Retest")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1RetestCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2RetestCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorRetestCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffRetestCount1++;
			        }
			    }
				//Clarification query
			    if (thirdColumnValue.contains("Clarification")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1ClarificationQueryCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2ClarificationQueryCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorClarificationQueryCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffClarificationQueryCount1++;
			        }
			    }
				//Running
			    if (thirdColumnValue.contains("Running")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1RunningCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2RunningCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorRunningCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffRunningCount1++;
			        }
			    }
			
			    // Bug Audit counts
			    if (thirdColumnValue.contains("Bug Audit")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1BugAuditCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2BugAuditCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorBugAuditCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffBugAuditCount1++;
			        }
			    }
			
			    // Query counts
			    if (thirdColumnValue.contains("Query ")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1QueryCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2QueryCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorQueryCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffQueryCount1++;
			        }
			    }
			
			    // Device Brick counts
			    if (thirdColumnValue.contains("Device Brick")) {
			        if (fourthColumnValue.contains("NC1")) {
			            nc1DeviceBrickCount1++;
			        } else if (fourthColumnValue.contains("NC2")) {
			            nc2DeviceBrickCount1++;
			        } else if (fourthColumnValue.contains("User Error")) {
			            userErrorDeviceBrickCount1++;
			        } else if (fourthColumnValue.contains("Waived")) {
			            waiveoffDeviceBrickCount1++;
			        }
			    }
			}
			    }
			}
	                        }
	                    }
	                }
	            }
			}
			  
			int totalExecutionAuditRow1 = nc1ExecutionAuditCount1 + nc2ExecutionAuditCount1 + userErrorExecutionAuditCount1 + waiveoffExecutionAuditCount1;
			int totalNonExecutionAuditRow1 = nc1NonExecutionAuditCount1 + nc2NonExecutionAuditCount1 + userErrorNonExecutionAuditCount1 + waiveoffNonExecutionAuditCount1;
			int totalParkRow1 = nc1ParkCount1 + nc2ParkCount1 + userErrorParkCount1 + waiveoffParkCount1;
			int totalTimeSheetRow1 = nc1TimeSheetCount1 + nc2TimeSheetCount1 + userErrorTimeSheetCount1  + waiveoffTimeSheetCount1;
			int totalBuildRow1 = nc1BuildCount1 + nc2BuildCount1 + userErrorBuildCount1 + waiveoffBuildCount1;
			int totalLogRow1 = nc1LogCount1 + nc2LogCount1 + userErrorLogCount1 + waiveoffLogCount1;
			int totalRetestRow1 = nc1RetestCount1 + nc2RetestCount1 + userErrorRetestCount1 + waiveoffRetestCount1;
			int totalClarificationQueryRow1 = nc1ClarificationQueryCount1 + nc2ClarificationQueryCount1 + userErrorClarificationQueryCount1 + waiveoffClarificationQueryCount1;
			int totalRunningRow1 = nc1RunningCount1 + nc2RunningCount1 + userErrorRunningCount1 + waiveoffRunningCount1;
			int totalBugAuditRow1 = nc1BugAuditCount1 + nc2BugAuditCount1 + userErrorBugAuditCount1 + waiveoffBugAuditCount1;
			int totalQueryRow1 = nc1QueryCount1 + nc2QueryCount1 + userErrorQueryCount1 + waiveoffQueryCount1;
			int totalDeviceBrickRow1 = nc1DeviceBrickCount1 + nc2DeviceBrickCount1 + userErrorDeviceBrickCount1 + waiveoffDeviceBrickCount1;
		
			
			int totalnc11 = nc1BugAuditCount1 + nc1DeviceBrickCount1 + nc1ExecutionAuditCount1 + nc1NonExecutionAuditCount1 + nc1QueryCount1 + nc1ParkCount1 + nc1ClarificationQueryCount1 + nc1RetestCount1 + nc1LogCount1 + nc1BuildCount1 + nc1TimeSheetCount1 + nc1RunningCount1;
			int totalnc21 = nc2BugAuditCount1 + nc2DeviceBrickCount1 + nc2ExecutionAuditCount1 + nc2NonExecutionAuditCount1 + nc2QueryCount1 + nc2ClarificationQueryCount1 + nc2RetestCount1 + nc2LogCount1 + nc2BuildCount1 + nc2TimeSheetCount1 + nc2ParkCount1 + nc2RunningCount1;
			int totaluserError1 = userErrorBugAuditCount1 + userErrorDeviceBrickCount1 + userErrorExecutionAuditCount1 + userErrorNonExecutionAuditCount1 + userErrorQueryCount1 + userErrorClarificationQueryCount1 + userErrorRetestCount1 + userErrorLogCount1 + userErrorBuildCount1 + userErrorTimeSheetCount1 + userErrorParkCount1 + userErrorRunningCount1;
			int totalwaiveoff1 = waiveoffExecutionAuditCount1 + waiveoffNonExecutionAuditCount1 + waiveoffBugAuditCount1 + waiveoffQueryCount1 + waiveoffDeviceBrickCount1 + waiveoffClarificationQueryCount1 + waiveoffRetestCount1 + waiveoffLogCount1 + waiveoffBuildCount1 + waiveoffTimeSheetCount1 + waiveoffParkCount1 + waiveoffRunningCount1;
			int total1 = totalExecutionAuditRow1 + totalNonExecutionAuditRow1 + totalBugAuditRow1 + totalQueryRow1 + totalDeviceBrickRow1 + totalClarificationQueryRow1 + totalRetestRow1 + totalLogRow1 + totalBuildRow1 + totalTimeSheetRow1 + totalParkRow1 + totalRunningRow1;
			
			// Execution Audit
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Execution Audit</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1ExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2ExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalExecutionAuditRow1).append("</td>");
			tableHtml.append("</tr>");

			// Non-Execution Audit
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Non - Execution Audit</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1NonExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2NonExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorNonExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffNonExecutionAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalNonExecutionAuditRow1).append("</td>");
			tableHtml.append("</tr>");
			
			//Park
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Park</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1ParkCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2ParkCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorParkCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffParkCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalParkRow1).append("</td>");
			tableHtml.append("</tr>");

			//Time sheet
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Time Sheet</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1TimeSheetCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2TimeSheetCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorTimeSheetCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffTimeSheetCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalTimeSheetRow1).append("</td>");
			tableHtml.append("</tr>");

			//Build
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Build</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1BuildCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2BuildCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorBuildCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffBuildCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalBuildRow1).append("</td>");
			tableHtml.append("</tr>");

			//Log
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Log</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1LogCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2LogCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorLogCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffLogCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalLogRow1).append("</td>");
			tableHtml.append("</tr>");

			//Retest
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Retest</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1RetestCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2RetestCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorRetestCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffRetestCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalRetestRow1).append("</td>");
			tableHtml.append("</tr>");

			//Clarification query
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Clarification Query</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1ClarificationQueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2ClarificationQueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorClarificationQueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffClarificationQueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalClarificationQueryRow1).append("</td>");
			tableHtml.append("</tr>");

			//Running
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Running</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1RunningCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2RunningCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorRunningCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffRunningCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalRunningRow1).append("</td>");
			tableHtml.append("</tr>");


			// Bug Audit
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Bug Audit</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1BugAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2BugAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorBugAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffBugAuditCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalBugAuditRow1).append("</td>");
			tableHtml.append("</tr>");

			// SLA Miss
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Query</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1QueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2QueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorQueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffQueryCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalQueryRow1).append("</td>");
			tableHtml.append("</tr>");

			// Device Brick
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber4++).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>Device Brick</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc1DeviceBrickCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(nc2DeviceBrickCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(userErrorDeviceBrickCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(waiveoffDeviceBrickCount1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalDeviceBrickRow1).append("</td>");
			tableHtml.append("</tr>");
			
			//Total
			tableHtml.append("<tr>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;background-color: lightblue;'>Grand Total </td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalnc11).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalnc21).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totaluserError1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(totalwaiveoff1).append("</td>");
			tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(total1).append("</td>");
			tableHtml.append("</tr>");
			
			tableHtml.append("</table><br>");
			
	      //Project wise clarification
			tableHtml.append("<p><br><p style='font-weight: bold;'>Project Wise Audit Count : </p>");
			tableHtml.append("<table border='1' style='text-align: center;'>");
			tableHtml.append("<tr style='background-color: lightblue;'><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold; background-color: lightblue;'>SL.NO</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Project</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Count</th></tr>");

			Map<String, Integer> projectCount = new HashMap<>();
			int serialNumber1 = 1;
			int rowCounter2 = 0;
			for (Row row : rowList) {
				if (row != null) {
					Cell createDateCell = row.getCell(1);
					if (createDateCell != null) {
						if (createDateCell.getCellType() == CellType.NUMERIC) {
							LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
							if (createDate.getYear() == lastMonth.getYear()  ){ {
								if (row != null) {
//									if (rowCounter2 > 0) { // Ignore the 1st row and start from the 2nd row
										Cell projectCell = row.getCell(7);
										if (projectCell != null) {
											String project = projectCell.getStringCellValue();
											projectCount.put(project, projectCount.getOrDefault(project, 0) + 1);
										}
									
									rowCounter2++;
								}
							}
							}
						}
					}
				}
			}

			int rowCounter1 = 0;
			for (Map.Entry<String, Integer> entry : projectCount.entrySet()) {
				tableHtml.append("<tr>");
				tableHtml.append("<td style='text-align: center;border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber1).append("</td>");
				tableHtml.append("<td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(entry.getKey()).append("</td><td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(entry.getValue()).append("</td></tr>");
				rowCounter1++;
				if (rowCounter1 > 0) {
					serialNumber1++;
				}
			}
			tableHtml.append("</table><br>");
			
			// Your existing HTML table setup
			tableHtml.append("<p style='font-weight: bold;'>Audits raised in the Last 7 days: " + "</p>");
			tableHtml.append("<table border='1'>");
			int serialNumber2 = 1;
			int rowCounter = 0;
			for (Row row : rowList) {
			    if (row != null) {
			        if (rowCounter == 0) { // Print the first row as it is
			            tableHtml.append("<tr style='background-color: lightblue;'>");
			            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>SL.NO</td>");
			            for (int j = 0; j < row.getLastCellNum(); j++) {
			                Cell cell = row.getCell(j);
			                if (j == 5 || j == 6) { // Skip 5th and 6th columns (index 4 and 6)
	                            continue;
	                        }
			                if (cell != null) {
			                    tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>").append(cell.getStringCellValue()).append("</td>");
			                    System.out.println(cell.getStringCellValue());
			                } else {
			                    tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>&nbsp;</td>");
			                }
			            }
			            tableHtml.append("</tr>");
			            rowCounter++;
			        } else { // Check if create date is in last month and print the row
			            Cell createDateCell = row.getCell(1); // Assuming the create date is in the 2nd column
			            if (createDateCell != null) {
			                LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
			                if (createDate.getYear() == lastMonth.getYear() ) {
//			                    if (rowCounter == 1) { // Check if it's the first data row
//			                        tableHtml.append("<tr style='background-color: lightblue;'>");
//			                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>(link unavailable)</td>");
//			                    } else {
			                        tableHtml.append("<tr>");
			                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber2).append("</td>");
			                    
			                    for (int j = 0; j < row.getLastCellNum(); j++) {
			                        if (j == 5 || j == 6) { // Skip 5th and 6th columns (index 4 and 6)
			                            continue;
			                        }
			                        Cell cell = row.getCell(j);
			                        if (cell != null) {
			                            String cellValue;
			                            switch (cell.getCellType()) {
			                                case STRING:
			                                    cellValue = cell.getStringCellValue();
			                                    if (j == 0 && rowCounter != 0) { // Convert 2nd column to link format starting from the 2nd row
			                                        String linkText = cellValue.substring(cellValue.indexOf("/issues/") + 8);
			                                        cellValue = "<a href='" + cellValue + "'>" + linkText + "</a>";
			                                    }
			                                    break;
			                                case NUMERIC:
			                                    if (DateUtil.isCellDateFormatted(cell)) {
			                                        double excelDate = cell.getNumericCellValue();
			                                        LocalDate date = LocalDate.ofEpochDay((long) (excelDate - 25569));
			                                        long daysBetween = java.time.temporal.ChronoUnit.DAYS.between(date, today);
			                                        cellValue = date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) + " - " + daysBetween + " days ago";
			                                        if (date.isBefore(fiveDaysAgo)) {
			                                            cellValue = cellValue + "</span>";
			                                        }
			                                    } else {
			                                        cellValue = String.valueOf(cell.getNumericCellValue());
			                                    }
			                                    break;
			                                case BOOLEAN:
			                                    cellValue = String.valueOf(cell.getBooleanCellValue());
			                                    break;
			                                default:
			                                    cellValue = "";
			                            }
			                            if (rowCounter == 1) { // Check if it's the first row
	                            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; '>").append(cellValue).append("</td>"); // Make the text bold and fill the cell with light blue color
	                        } else {
	                            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(cellValue).append("</td>");
	                        }
	                    } else {
	                        if (rowCounter == 0) { // Check if it's the first row
	                            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>&nbsp;</td>"); // Fill the cell with light blue color and make the text bold
	                        } else {
	                            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>&nbsp;</td>");
	                        }
	                    }
	                }
	                tableHtml.append("</tr>");
	                rowCounter++;
	                if (rowCounter > 1) { // Start numbering from 2nd row
	                    serialNumber2++;
	                }
	            }
	        }
			        }
			    }
			}
			
	        tableHtml.append("</table><br><br><br><p>This is an automated mail, if any clarifications please reach out to @rbharatm.</p>");
	        // Set the email body
	        String tableHtmlString = tableHtml.toString().replace("\n", "\\n");
	        ((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = `" + tableHtmlString + "`;", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
	        
	        // Rename the excel file
	        String newFileName = "AUDIT_REPORT_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xls";
	        File newFile = new File(downloadDir, newFileName);
	        if (!latestFile.renameTo(newFile)) {
	            throw new IOException("Failed to rename the file.");
	        }
	        String newdownloadedFilePath = newFile.getAbsolutePath();
	        
	        Thread.sleep(5000);
	        // Attach the file to the email
	        WebElement fileInput = driver.findElement(By.xpath("//input[@type='file']"));
	        fileInput.sendKeys(newdownloadedFilePath);
	        Thread.sleep(10000);
	        
	        // Send the email
	        // driver.findElements(By.xpath("//*[@aria-label='Send']")).get(0).click();
	        Thread.sleep(10000);
	        
	        // Cleanup
	        //driver.quit();
	    }
	}


