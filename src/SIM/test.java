package SIM;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hslf.record.Document;
import org.apache.poi.hssf.usermodel.*;
//import org.apache.poi.HSSF.usermodel.*;
import org.jsoup.Jsoup;
import org.openqa.selenium.By;
//import (link unavailable);
import org.openqa.selenium.InvalidArgumentException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Clarification {
    public static void main(String[] args) throws InterruptedException, IOException, AWTException {
        WebDriver driver = new FirefoxDriver();
        driver.manage().window().maximize();
//		driver.get("https://issues.amazon.com/issues/search?q=status%3A(Open)+containingFolder%3A(2857712e-96fa-4b2c-8a01-0328fe175eb0)+-createDate%3A(%5BNOW-2DAYS..NOW%5D)&sort=lastUpdatedDate+desc&selectedDocument=148be64f-ac5d-41fd-a9ee-366e1186a33e");	
//		//Enter username in sendkeys
//		driver.findElement(By.id("user_name_field")).sendKeys("rbharatm");
//		driver.findElement(By.id("user_name_btn")).click();
//		Thread.sleep(3000);
//		//Enter password in sendkeys
//		driver.findElement(By.id("password_field")).sendKeys("Bh@rath2008");
//		driver.findElement(By.id("password_btn")).click();
//		//driver.findElement(By.id("user_name")).sendKeys("");
//		//driver.findElement(By.id("password")).sendKeys("");
//		//driver.findElement(By.id("verify_btn")).click();
//		Thread.sleep(40000);
//		driver.findElement(By.xpath("//button[@class='btn btn-small dropdown-toggle']")).click();
//		driver.findElement(By.xpath("//a[@class='export-search-results']")).click();
//		driver.findElement(By.id("submit-custom-export-job")).click();
//		Thread.sleep(30000);
//		driver.findElement(By.xpath("//a[@data-link='html{>fileName} href{:~getAttachmentURL(~jobId, id, stack)}']")).click();
//	
//		driver.findElement(By.cssSelector("body")).sendKeys(Keys.CONTROL + "t");

        driver.get("https://ballard.amazon.com/owa/#path=/mail");
        Thread.sleep(5000);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        driver.findElement(By.xpath("//*[@title='Write a new message (N)']")).click();
        Thread.sleep(30000);
        //Enter the userid and groups ex:rbharatm@amazon.com
        driver.findElements(By.xpath("//*[@aria-label='To']")).get(0).sendKeys(" rbharatm@amazon.com");
        //Enter the subject to be entered in the mail.
        driver.findElements(By.xpath("//*[@placeholder='Add a subject']")).get(0).sendKeys("SLA MISS for clarification");

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
        StringBuilder tableHtml = new StringBuilder("<p>Hi All,<br><br>Please find the data below for the missed audit:</p><br>The date highlighted with yellow is issues more than 5 days from today and the userid highlighted is the action should be taken by the person requested the clarification.<br><br><table border='1'>");
        
        LocalDate today = LocalDate.now();
        LocalDate fiveDaysAgo = today.minusDays(5);

        
        // Summary section
        tableHtml.append("<p style='font-size: 18px; font-weight: bold;'>Summary:</p><br>");
        tableHtml.append("<p style='font-weight: bold;'>Daily count of clarifications : ");

        // Count of create date 2 days ago from today
        int clarificationCount = 0;
        for (Row row : rowList) {
            if (row != null) {
                Cell createDateCell = row.getCell(1); // assuming create date is in second column
                if (createDateCell != null) {
                    if (createDateCell.getCellType() == CellType.NUMERIC) {
                        LocalDate createDate = LocalDate.ofEpochDay((long) (createDateCell.getNumericCellValue() - 25569));
                        LocalDate twoDaysAgo = today.minusDays(2);
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
        tableHtml.append("<table border='1'>");
        tableHtml.append("<tr><th>Assignee</th><th>Count</th></tr>");

        Map<String, Integer> assigneeCount = new HashMap<>();
        for (Row row : rowList) {
            if (row != null) {
                Cell assigneeCell = row.getCell(row.getLastCellNum() - 1); // assuming assignee is in last column
                if (assigneeCell != null) {
                    String assignee = assigneeCell.getStringCellValue();
                    if (Arrays.asList("arajikan", "hsethura", "irfahmd", "kavipria", "nandhnr", "pawansaj", "perusri", "ppryankv", "qasadana", "qelavara", "rajamor", "rdkavith", "ssivs", "ssuryaku", "vithivit", "yaruravi").contains(assignee)) {
                        assigneeCount.put(assignee, assigneeCount.getOrDefault(assignee, 0) + 1);
                    }
                }
            }
        }

        for (Map.Entry<String, Integer> entry : assigneeCount.entrySet()) {
            tableHtml.append("<tr><td>").append(entry.getKey()).append("</td><td>").append(entry.getValue()).append("</td></tr>");
        }

        tableHtml.append("</table>");


       
        
     // Check for Labels Missing
    boolean labelsMissing = false;
    tableHtml.append("<p><br><p style='font-weight: bold;'>Labels Missing for : </p>");
    tableHtml.append("<table border='1'>");
    tableHtml.append("<tr><th>Submitter Identity</th><th>Issue ID</th></tr>");

    for (Row row : rowList) {
        if (row != null) {
            Cell labelsCell = row.getCell(3); // assuming Labels column is at index 3
            if (labelsCell == null || labelsCell.getStringCellValue() == null || labelsCell.getStringCellValue().isEmpty()) {
                labelsMissing = true;
                Cell submitterCell = row.getCell(4); // assuming Submitter Identity column is at index 5
                Cell issueIdCell = row.getCell(0); // assuming Issue ID is in the first column
                if (submitterCell != null && issueIdCell != null) {
                    tableHtml.append("<tr><td>").append(submitterCell.getStringCellValue()).append("</td><td>").append(issueIdCell.getStringCellValue()).append("</td></tr>");
                }
            }
        }
    }

    tableHtml.append("</table>");

        
        tableHtml.append("</tr><br><br>");
       
        for (Row row : rowList) {
            if (row != null) {
                tableHtml.append("<tr>");
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        String cellValue;
                        switch (cell.getCellType()) {
                            case STRING:
                                cellValue = cell.getStringCellValue();
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
                        if (row.getRowNum() == 0) { // first row
                            tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; font-weight: bold;'>").append(cellValue).append("</td>");
                        } else {
                            if (j == row.getLastCellNum() - 1) { // last column
                                if (row.getLastCellNum() > 1) {
                                    Cell lastCell = row.getCell(row.getLastCellNum() - 1);
                                    Cell secondLastCell = row.getCell(row.getLastCellNum() - 2);
                                    if (lastCell.getStringCellValue().equals(secondLastCell.getStringCellValue())) {
                                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px; background-color: yellow;'>").append(cellValue).append("</td>");
                                    } else {
                                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(cellValue).append("</td>");
                                    }
                                } else {
                                    tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(cellValue).append("</td>");
                                }
                            } else {
                                tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(cellValue).append("</td>");
                            }
                        }
                    } else {
                        tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>&nbsp;</td>");
                    }
                }
                tableHtml.append("</tr>");
            }
        }


        tableHtml.append("</table><br><br>");

       


        tableHtml.append("<p><br><br>Thanks<br>Bharath M.</p>");

        
        // Set the email body
        String tableHtmlString = tableHtml.toString().replace("\n", "\\n");
        ((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = " + tableHtmlString + ";", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
        
        // Rename the excel file
        String newFileName = "SLA_Miss_Report_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xls";
        File newFile = new File(downloadDir, newFileName);
        if (!latestFile.renameTo(newFile)) {
            throw new IOException("Failed to rename the file.");
        }
        String newdownloadedFilePath = newFile.getAbsolutePath();
        
        // Attach the file to the email
        Thread.sleep(5000);
        WebElement fileInput = driver.findElement(By.xpath("//input[@type='file']"));
        fileInput.sendKeys(newdownloadedFilePath);
        Thread.sleep(10000);
        
        // Send the email
         //driver.findElements(By.xpath("//*[@aria-label='Send']")).get(0).click();
        Thread.sleep(10000);
//      driver.close();
    }
}




for (Row row : rowList) {
    if (row != null) {
        tableHtml.append("<tr>");
        for (int j = 0; j < row.getLastCellNum(); j++) {
            Cell cell = row.getCell(j);
            if (cell != null) {
                String cellValue;
                switch (cell.getCellType()) {
                    case STRING:
                        cellValue = cell.getStringCellValue();
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
                tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>").append(cellValue).append("</td>");
            } else {
                tableHtml.append("<td style='text-align: center; border: 2px solid black; padding: 5px;'>&nbsp;</td>");
            }
        }

        tableHtml.append("</tr>");
    }
}
tableHtml.append("</table><br><br><br><p>Thanks<br>Bharath M.</p>");


make few changes for the first row