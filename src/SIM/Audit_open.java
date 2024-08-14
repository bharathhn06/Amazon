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

public class Audit_open {

	 public static void main(String[] args) throws InterruptedException, IOException, AWTException {
	        WebDriver driver = new FirefoxDriver();
	        driver.manage().window().maximize();
			driver.get("https://issues.amazon.com/issues/search?q=containingFolder%3A(2a42d4ee-18c1-4cf8-8a1e-b1344ff622c8)+status%3A(Open)+createDate%3A(%5BNOW-7DAYS..NOW%5D)&sort=lastUpdatedDate+desc&selectedDocument=56d9e243-48e7-4694-b20b-bdfd3d1a98bb");	
			//Enter username in sendkeys
			driver.findElement(By.id("user_name_field")).sendKeys("rbharatm");
			driver.findElement(By.id("user_name_btn")).click();
			Thread.sleep(3000);
			//Enter password in sendkeys
			driver.findElement(By.id("password_field")).sendKeys("Bh@rath2008");
			driver.findElement(By.id("password_btn")).click();
			//driver.findElement(By.id("user_name")).sendKeys("");
			//driver.findElement(By.id("password")).sendKeys("");
			//driver.findElement(By.id("verify_btn")).click();
			Thread.sleep(40000);
			driver.findElement(By.xpath("//button[@class='btn btn-small dropdown-toggle']")).click();
			driver.findElement(By.xpath("//a[@class='export-search-results']")).click();
			driver.findElement(By.id("submit-custom-export-job")).click();
			Thread.sleep(30000);
			driver.findElement(By.xpath("//a[@data-link='html{>fileName} href{:~getAttachmentURL(~jobId, id, stack)}']")).click();
		
			driver.findElement(By.cssSelector("body")).sendKeys(Keys.CONTROL + "t");

	        driver.get("https://ballard.amazon.com/owa/#path=/mail");
	        Thread.sleep(5000);
	        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
	        driver.findElement(By.xpath("//*[@title='Write a new message (N)']")).click();
	        Thread.sleep(30000);
	        //Enter the userid and groups ex:rbharatm@amazon.com
	        driver.findElements(By.xpath("//*[@aria-label='To']")).get(0).sendKeys("tdhivya@amazon.com ; rbharatm@amazon.com");
	        //Enter the subject to be entered in the mail.
	        driver.findElements(By.xpath("//*[@placeholder='Add a subject']")).get(0).sendKeys("SLA MISS for Execution and Non-Execution Audit");


	        // Find the latest downloaded file
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
	        StringBuilder tableHtml = new StringBuilder("<p>Hi All," + "<br><br>" + "Please find the data below for the missed audit:</p>" + "<br>" + "The date highlighted with yellow is issues more than 5 days from today and the userid highlighted is the action should be taken by the person requested the clarification." + "<br><br>");

	        LocalDate today = LocalDate.now();
	        LocalDate fiveDaysAgo = today.minusDays(5);

	        // Summary section
	        tableHtml.append("<p style='font-size: 18px; font-weight: bold;'>Summary:</p><br>");
	        int clarificationCount = 0;
	        LocalDate Today = LocalDate.now();
	        LocalDate lastWeek = Today.minusWeeks(4);
	        for (Row row : rowList) {
	            if (row != null) {
	                Cell createDateCell = row.getCell(1);
	                if (createDateCell != null) {
	                    if (createDateCell.getCellType() == CellType.NUMERIC) {
	                        LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
	                        if (!createDate.isBefore(lastWeek)) {
	                            clarificationCount++;
	                        }
	                    }
	                }
	            }
	        }

	        tableHtml.append("<p style='font-weight: bold;'>Total No.of Audit for the past week : " + clarificationCount + "</p>");


//	        // SDA tagged issues count
//	        tableHtml.append("<p><br><p style='font-weight: bold;'>SDA tagged issues : </p>");
//	        tableHtml.append("<table border='1' style='text-align: center;'>");
//	        tableHtml.append("<tr style='background-color: lightblue;'><th style='text-align: center; font-weight: bold;'>Assignee</th><th style='text-align: center; font-weight: bold;'>Count</th></tr>");
//	        Map<String, Integer> assigneeCount = new HashMap<>();
//	        for (Row row : rowList) {
//	            if (row != null) {
//	                Cell assigneeCell = row.getCell(row.getLastCellNum() - 1);
//	                if (assigneeCell != null) {
//	                    String assignee = assigneeCell.getStringCellValue();
//	                    if (Arrays.asList("arajikan", "hsethura", "irfahmd", "kavipria", "nandhnr", "pawansaj", "perusri", "ppryankv", "qasadana", "qelavara", "rajamor", "rdkavith", "ssivs", "ssuryaku", "vithivit", "yaruravi").contains(assignee)) {
//	                        assigneeCount.put(assignee, assigneeCount.getOrDefault(assignee, 0) + 1);
//	                    }
//	                }
//	            }
//	        }
//	        for (Map.Entry<String, Integer> entry : assigneeCount.entrySet()) {
//	            tableHtml.append("<tr><td style='text-align: center;'>").append(entry.getKey()).append("</td><td style='text-align: center;'>").append(entry.getValue()).append("</td></tr>");
//	        }
//	        tableHtml.append("</table>");
//
//	        // Check for Labels Missing
//	        boolean labelsMissing = false;
//	        tableHtml.append("<p><br><p style='font-weight: bold;'>Labels Missing : </p>");
//	        tableHtml.append("<table border='1' style='text-align: center;'>");
//	        tableHtml.append("<tr style='background-color: lightblue;'><th style='text-align: center; font-weight: bold;'>Issue ID</th><th style='text-align: center; font-weight: bold;'>Submitter Identity</th><th style='text-align: center; font-weight: bold;'>Assignee Identity</th></tr>");
//	        for (Row row : rowList) {
//	            if (row != null) {
//	                Cell labelsCell = row.getCell(3);
//	                if (labelsCell == null || labelsCell.getStringCellValue() == null || labelsCell.getStringCellValue().isEmpty()) {
//	                    labelsMissing = true;
//	                    Cell issueIdCell = row.getCell(0);
//	                    Cell submitterCell = row.getCell(4);
//	                    Cell assigneeCell = row.getCell(5);
//	                    if (issueIdCell != null && submitterCell != null && assigneeCell != null) {
//	                        tableHtml.append("<tr><td style='text-align: center;'>").append(issueIdCell.getStringCellValue()).append("</td><td style='text-align: center;'>").append(submitterCell.getStringCellValue()).append("</td><td style='text-align: center;'>").append(assigneeCell.getStringCellValue()).append("</td></tr>");
//	                    }
//	                }
//	            }
//	        }
	        tableHtml.append("</table><br>");

	      
	        tableHtml.append("<table border='1'>");
	        int serialNumber = 1;
	        int rowCounter = 0;
	        for (Row row : rowList) {
	            if (row != null) {
	                if (rowCounter == 0) { // Check if it's the first row
	                    tableHtml.append("<tr style='background-color: lightblue;'>"); // Set the background color to light blue
	                } else {
	                    tableHtml.append("<tr>");
	                }
	                
	                    if (rowCounter == 0) { // Check if it's the first row
	                        tableHtml.append("<tr style='background-color: lightblue;'>"); // Set the background color to the specified color
	                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>SL.NO</td>"); // Add "SL.NO" to the first cell
	                    } else {
	                        tableHtml.append("<tr>");
	                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber).append("</td>"); // Add serial number cell
	                    }
	                for (int j = 0; j < row.getLastCellNum(); j++) {
	                	if ( j == 4 || j == 5) { // Skip 5th and 6th columns (index 4 and 6)
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
                                    if (j == row.getLastCellNum() - 1 && j > 0) { // Check if it's the last column and not the first column
                                        Cell previousCell = row.getCell(j - 1);
                                        if (previousCell != null && previousCell.getStringCellValue().equals(cellValue)) {
                                            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: yellow;'>").append(cellValue).append("</td>"); // Highlight the last column with yellow color
                                            continue;
                                        }
                                    }
                                    break;
	                            case NUMERIC:
	                                if (DateUtil.isCellDateFormatted(cell)) {
	                                    double excelDate = cell.getNumericCellValue();
	                                    LocalDate date = LocalDate.ofEpochDay((long) (excelDate - 25569));
	                                    long daysBetween = java.time.temporal.ChronoUnit.DAYS.between(date, today);
	                                    cellValue = date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) + " - " + daysBetween + " days ago";
	                                    if (date.isBefore(fiveDaysAgo)) {
	                                        cellValue =  cellValue + "</span>";
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
	                        if (rowCounter == 0) { // Check if it's the first row
	                            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>").append(cellValue).append("</td>"); // Make the text bold and fill the cell with light blue color
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
                        serialNumber++;
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

