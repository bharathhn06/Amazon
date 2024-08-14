package SIM;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Today;
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

public class sla_miss_clarification {
    public static void main(String[] args) throws InterruptedException, IOException, AWTException {
        WebDriver driver = new FirefoxDriver();
        driver.manage().window().maximize();
		driver.get("https://issues.amazon.com/issues/search?q=status%3A(Open)+containingFolder%3A(2857712e-96fa-4b2c-8a01-0328fe175eb0)+-createDate%3A(%5BNOW-0DAYS..NOW%5D)&sort=lastUpdatedDate+desc&selectedDocument=148be64f-ac5d-41fd-a9ee-366e1186a33e");	
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

        // Compose a new email
        driver.findElement(By.xpath("//*[@title='Write a new message (N)']")).click();
        Thread.sleep(40000);

        // Enter the user IDs and groups
        driver.findElements(By.xpath("//*[@aria-label='To']")).get(0).sendKeys("ereader-ds@amazon.com; rbharatm@amazon.com");

        // Enter the subject of the mail
        driver.findElements(By.xpath("//*[@placeholder='Add a subject']")).get(0).sendKeys("SLA MISS for clarification " + LocalDate.now() );

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

        // Build email content
        StringBuilder tableHtml = new StringBuilder(
        	    "<p>Hi Team,<br><br>" +
        	    "Please find below the Clarification Summary.</p><br>" +
        	    "<span style='background-color: yellow;'>Ask from Teams:</span><br><br>");
        tableHtml.append("<ul>")
                 .append("<li>Clarifications must be resolved within 2 days of SLA.</li>")
                 .append("<li>Daily Audit will be done and deviation will be tagged for SIMs with more than 2 days SLA.</li>")
                 .append("<li>SIM with any dependency and added with 'has-dependency' label should also be resolved by max of 4 days. Component owners should take ownership, track the dependency with QAE and close the SIM before the milestone signoff date.</li>")
                 .append("</ul>");
        LocalDate today = LocalDate.now();
        LocalDate fiveDaysAgo = today.minusDays(5);
        
        
     // Summary section
        tableHtml.append("<p style='font-size: 18px; font-weight: bold;'>Summary:</p><br>");
        tableHtml.append("<p style='font-weight: bold;'>Total Clarifications created on " + today + " : ");
        // Count of create date 2 days ago from today
        int clarificationCount = 0;
        for (Row row : rowList) {
            if (row != null) {
                Cell createDateCell = row.getCell(1); 
                if (createDateCell != null) {
                    if (createDateCell.getCellType() == CellType.NUMERIC) {
                        LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
                        LocalDate twoDaysAgo = today.minusDays(0);
                        if (createDate.equals(twoDaysAgo)) {
                            clarificationCount++;
                        }
                    }
                }
            }
         
        }
        tableHtml.append(clarificationCount);
        tableHtml.append("</p>");

        // SDA tagged issues count
		
		tableHtml.append("<p><br><p style='font-weight: bold;'>SDA tagged issues : </p>");
		tableHtml.append("<table border='1' style='text-align: center;'>");
		tableHtml.append("<tr style='background-color: lightblue;'><th style='text-align: center; border: 2px solid black; padding: 5px; font-weight: bold; background-color: lightblue;'>SL.NO</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Assignee</th><th style='text-align: center; border: 2px solid black; padding: 5px; font-weight: bold;'>Count</th></tr>");
		Map<String, Integer> assigneeCount = new HashMap<>();
		int serialNumber2 = 1; 
		for (Row row : rowList) {
		    if (row != null) {
		        Cell assigneeCell = row.getCell(row.getLastCellNum() - 1);
		        Cell dateCell = row.getCell(1);
		        if (assigneeCell != null && dateCell != null) {
		            if (dateCell.getCellType() == CellType.NUMERIC) {
		                String assignee = assigneeCell.getStringCellValue();
		                LocalDate date = LocalDate.ofEpochDay((long) (dateCell.getNumericCellValue() - 25569));
		                if (!date.isEqual(today) && !date.isEqual(today.minusDays(1))) {
		                    if (Arrays.asList("psubashi", "ghaiunna", "sheelj", "dprincy", "hverames", "irfahmd", "karrrame", "varunkri", "sabsanga", "aamidhun", "llsrut", "sunilkjy", "atnirmal", "muthumzs", "avignest", "dprajwkv", "ssuryaku", "fssayeed", "akkigopi", "arajikan", "qasadana", "perusrip", "qelavara", "yaruravi", "hsethura", "kavipria", "pawansaj", "vithivit", "nandhnr", "pryankv", "rajamor", "rdkavith", "ssivs", "rbharatm", "kalmeshw", "rajeeshb", "rsprathi", "nrnjnvj", "shubhrbe", "ssikl", "redklthl").contains(assignee)) {
		                        assigneeCount.put(assignee, assigneeCount.getOrDefault(assignee, 0) + 1);
		                    }
		                }
		            } else {
		                System.out.println("Date cell is not numeric");
		            }
		        }
		    }
		}
		for (Map.Entry<String, Integer> entry : assigneeCount.entrySet()) {
		    tableHtml.append("<tr><td style='text-align: center;border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber2).append("</td><td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(entry.getKey()).append("</td><td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(entry.getValue()).append("</td></tr>");
		    serialNumber2++; 
		}

        tableHtml.append("</table>");


        // Check for Labels Missing
        boolean labelsMissing = false;
        int rowCounter = 0;
        tableHtml.append("<p><br><p style='font-weight: bold;'>Labels Missing : </p>");
        tableHtml.append("<table border='1' style='text-align: center;'>");
        tableHtml.append("<tr style='background-color: lightblue;'><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Issue ID</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Submitter Identity</th><th style='text-align: center; border: 2px solid black; padding: 5px; font-weight: bold;'>Assignee Identity</th></tr>");
        for (Row row : rowList) {
            if (row != null) {
                Cell labelsCell = row.getCell(5);
                Cell dateCell = row.getCell(1);
                if (labelsCell == null || labelsCell.getStringCellValue() == null || labelsCell.getStringCellValue().isEmpty()) {
                    if (dateCell != null && dateCell.getCellType() == CellType.NUMERIC) {
                        LocalDate date = LocalDate.ofEpochDay((long) (dateCell.getNumericCellValue() - 25569));
                        if (!date.isEqual(today) && !date.isEqual(today.minusDays(1))) {
                            labelsMissing = true;
                            Cell issueIdCell = row.getCell(0);
                            Cell submitterCell = row.getCell(8); 
                            Cell assigneeCell = row.getCell(9); 
                            if (issueIdCell != null && submitterCell != null && assigneeCell != null) {
                                String issueIdValue = issueIdCell.getStringCellValue();
                                if (rowCounter == 0 && issueIdValue.contains("/issues/")) { // Convert 2nd column to link format starting from the 2nd row
                                    String linkText = issueIdValue.substring(issueIdValue.indexOf("/issues/") + 8);
                                    issueIdValue = "<a href='" + issueIdValue + "'>" + linkText + "</a>";
                                }
                                tableHtml.append("<tr><td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(issueIdValue).append("</td><td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(submitterCell.getStringCellValue()).append("</td><td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(assigneeCell.getStringCellValue()).append("</td></tr>");
                            }
                        }
                    } else {
                        System.out.println("Date cell is not numeric");
                    }
                }
            }
        }
        
        
        tableHtml.append("</table>");
        //Project wise clarification
        tableHtml.append("<p><br><p style='font-weight: bold;'>Project Wise Clarification raised on " + today +": </p>");
        tableHtml.append("<table border='1' style='text-align: center;'>");
        tableHtml.append("<tr style='background-color: lightblue;'><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold; background-color: lightblue;'>SL.NO</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Project</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Count</th></tr>");
        Map<String, Integer> projectCount = new HashMap<>();
        int serialNumber1 = 1; 
        for (Row row : rowList) {
            
        	
              if (row != null) {
                  Cell createDateCell = row.getCell(1);
                  Cell projectCell = row.getCell(7);
                  if (createDateCell != null && projectCell != null) {
                      if (createDateCell.getCellType() == CellType.NUMERIC) {
                          LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
                          LocalDate twoDaysAgo = today.minusDays(0);
                          if (createDate.equals(twoDaysAgo)) {
                              String project = projectCell.getStringCellValue();
                              projectCount.put(project, projectCount.getOrDefault(project, 0) + 1);
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
        tableHtml.append("</table>");
        tableHtml.append("<br>");
        
//        tableHtml.append("<p><br><p style='font-weight: bold;'>Project Wise Clarification raised today : </p>");
//        tableHtml.append("<table border='1' style='text-align: center;'>");
//        tableHtml.append("<tr style='background-color: lightblue;'><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>SL.NO</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Project</th><th style='text-align: center;border: 2px solid black; padding: 5px; font-weight: bold;'>Count</th></tr>");
//
//        Map<String, Integer> projectCount = new HashMap<>();
//        int serialNumber1 = 1;
//
//        for (Row row : rowList) {
//            if (row != null) {
//                Cell createDateCell = row.getCell(1);
//                Cell projectCell = row.getCell(7);
//                if (createDateCell != null && projectCell != null) {
//                    if (createDateCell.getCellType() == CellType.NUMERIC) {
//                        LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
//                        LocalDate twoDaysAgo = today.minusDays(0);
//                        if (createDate.equals(twoDaysAgo)) {
//                            String project = projectCell.getStringCellValue();
//                            projectCount.put(project, projectCount.getOrDefault(project, 0) + 1);
//                        }
//                    }
//                }
//            }
//        }
//
//        int rowCounter1 = 0;
//        for (Map.Entry<String, Integer> entry : projectCount.entrySet()) {
//            
//                tableHtml.append("<tr>");
//                tableHtml.append("<td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(serialNumber1).append("</td>");
//            
//            tableHtml.append("<td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(entry.getKey()).append("</td><td style='text-align: center;border: 2px solid black; padding: 5px;'>").append(entry.getValue()).append("</td></tr>");
//            rowCounter1++;
//            if (rowCounter1 > 0) {
//                serialNumber1++;
//            }
//        }
//
//        tableHtml.append("</table>");
//        tableHtml.append("<br>");



        tableHtml.append("</table>");
        tableHtml.append("<br>");
        tableHtml.append("<p><br><p style='font-weight: bold;'>Open Clarifications with SLA : </p>");
        tableHtml.append("<table border='1'>");
        int serialNumber = 1; // Initialize the serial number
        for (Row row : rowList) {
            if (row != null) {
                boolean printRow = true;
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    if (j == 3 || j == 4 || j == 6) { // Skip 5th and 7th columns (index 4 and 6)
                        continue;
                    }
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                            double excelDate = cell.getNumericCellValue();
                            LocalDate date = LocalDate.ofEpochDay((long) (excelDate - 25569));
                            if (date.isAfter(today.minusDays(2)) && date.isBefore(today.plusDays(1))) { // Check if date is within the last 2 days
                                printRow = false;
                                break;
                            }
                        }
                    }
                }
                if (printRow) {
                    if (rowCounter == 0) { // Check if it's the first row
                        tableHtml.append("<tr style='background-color: lightblue;'>"); // Set the background color to the specified color
                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>SL.NO</td>"); // Add "SL.NO" to the first cell
                    } else {
                        tableHtml.append("<tr>");
                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue;'>").append(serialNumber).append("</td>"); // Add serial number cell
                    }
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        if (j == 3 || j == 4 || j == 6) { // Skip 5th and 7th columns (index 4 and 6)
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
                                            cellValue = "<span style='background-color: yellow; color: red;'>" + cellValue + "</span>";
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
                                
                                    tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>").append(cellValue).append("</td>"); // Make the text bold and fill the cell with the specified color
                                
                            } else {
                                tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(cellValue).append("</td>");
                            }
                        } else {
                            if (rowCounter == 0) { // Check if it's the first row
                                tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: lightblue; font-weight: bold;'>&nbsp;</td>"); // Fill the cell with the specified color and make the text bold
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
        }
        tableHtml.append("</table><br><br><br><br><p>This is an automated mail, if any clarifications please reach out to @rbharatm.</p>"); // Set the email body
        String tableHtmlString = tableHtml.toString().replace("\n", "\\n");
        ((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = `" + tableHtmlString + "`;", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
        
        // Rename the excel file
        String newFileName = "SLA_Miss_Report_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xls";
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
