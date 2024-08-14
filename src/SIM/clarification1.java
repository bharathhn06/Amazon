package SIM;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;

public class clarification1 {
    public static void main(String[] args) throws InterruptedException, IOException, AWTException {
        WebDriver driver = new FirefoxDriver();
        driver.manage().window().maximize();
        driver.get("https://ballard.amazon.com/owa/#path=/mail");
        Thread.sleep(5000);

        // Compose a new email
        driver.findElement(By.xpath("//*[@title='Write a new message (N)']")).click();
        Thread.sleep(30000);

        // Enter the user IDs and groups
        driver.findElements(By.xpath("//*[@aria-label='To']")).get(0).sendKeys("tdhivya@amazon.com; rbharatm@amazon.com");

        // Enter the subject of the mail
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
        
        // Set the email body
        String tableHtmlString = tableHtml.toString().replace("\n", "\\n");
        ((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = `" + tableHtmlString + "`;", driver.findElement(By.xpath("//*[@aria-label='Message body']")));
        
        // Rename the excel file
        String newFileName = "SLA_Miss_Report_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xls";
        File newFile = new File(downloadDir, newFileName);
        if (!latestFile.renameTo(newFile)) {
            throw new IOException("Failed to rename the file.");
        }
        String newdownloadedFilePath = newFile.getAbsolutePath();
        
        // Attach the file to the email
        WebElement fileInput = driver.findElement(By.xpath("//input[@type='file']"));
        fileInput.sendKeys(newdownloadedFilePath);
        Thread.sleep(10000);
        
        // Send the email
        // driver.findElements(By.xpath("//*[@aria-label='Send']")).get(0).click();
        Thread.sleep(10000);
        
        // Cleanup
        driver.quit();
    }
}
