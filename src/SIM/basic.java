package SIM;

import java.awt.AWTException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class basic {
    public static void main(String[] args) throws AWTException {
        WebDriver driver = new FirefoxDriver();
        try {
            driver.manage().window().maximize();
            driver.get("https://ballard.amazon.com/owa/#path=/mail");
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));
            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@title='Write a new message (N)']"))).click();
            Thread.sleep(10000);
            // Wait for the email fields to be visible
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@aria-label='To']"))).sendKeys("tdhivya@amazon.com ; rbharatm@amazon.com");
            driver.findElement(By.xpath("//*[@placeholder='Add a subject']")).sendKeys("SLA MISS for clarification");

            // Locate the latest downloaded file
            File downloadDir = new File("C:\\Users\\rbharatm\\Downloads");
            File[] files = downloadDir.listFiles();
            File latestFile = null;
            long latestTimestamp = 0;
            for (File file : files) {
                if (file.lastModified() > latestTimestamp) {
                    latestFile = file;
                    latestTimestamp = file.lastModified();
                }
            }

            if (latestFile != null) {
                String downloadedFilePath = latestFile.getAbsolutePath();

                try (FileInputStream inputstream = new FileInputStream(downloadedFilePath);
                     HSSFWorkbook workbook = new HSSFWorkbook(inputstream)) {

                    HSSFSheet sheet = workbook.getSheetAt(0);
                    int rows = sheet.getLastRowNum();

                    StringBuilder tableHtml = new StringBuilder("<p>Hi All,<br><br>Please find the data below for the missed audit:</p><br>");
                    tableHtml.append("<p>The date highlighted with yellow is issues more than 5 days from today and the userid highlighted is the action should be taken by the person requested the clarification.</p><br>");
                    tableHtml.append("<table border='1'>");

                    LocalDate today = LocalDate.now();
                    LocalDate fiveDaysAgo = today.minusDays(5);

                    for (int i = 0; i <= rows; i++) {
                        HSSFRow row = sheet.getRow(i);
                        if (row != null) {
                            tableHtml.append("<tr>");
                            for (int j = 0; j < row.getLastCellNum(); j++) {
                                HSSFCell cell = row.getCell(j);
                                String style = (i == 0) ? "font-weight: bold; text-align: center; border: 2px solid black; padding: 5px;" : "text-align: center; border: 2px solid black; padding: 5px;";

                                if (cell != null) {
                                    tableHtml.append("<td style='").append(style).append("'>");
                                    switch (cell.getCellType()) {
                                        case STRING:
                                            if (j == row.getLastCellNum() - 1 && i > 0) {
                                                String currentValue = cell.getStringCellValue();
                                                boolean highlight = false;
                                                for (int k = 3; k < row.getLastCellNum() - 1; k++) {
                                                    if (sheet.getRow(i).getCell(k).getStringCellValue().equals(currentValue)) {
                                                        highlight = true;
                                                        break;
                                                    }
                                                }
                                                tableHtml.append(highlight ? "<span style='background-color: yellow; color: red;'>" + currentValue + "</span>" : currentValue);
                                            } else {
                                                tableHtml.append(cell.getStringCellValue());
                                            }
                                            break;
                                        case NUMERIC:
                                            double excelDate = cell.getNumericCellValue();
                                            LocalDateTime dateTime = LocalDateTime.ofEpochSecond((long) (excelDate - 25569) * 86400, 0, ZoneOffset.UTC);
                                            LocalDate date = dateTime.toLocalDate();
                                            tableHtml.append(date.isBefore(fiveDaysAgo) ? "<span style='background-color: yellow; color: red;'>" + dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) + "</span>" : dateTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
                                            break;
                                        case BOOLEAN:
                                            tableHtml.append(cell.getBooleanCellValue());
                                            break;
                                    }
                                    tableHtml.append("</td>");
                                } else {
                                    tableHtml.append("<td style='").append(style).append("'>&nbsp;</td>");
                                }
                            }
                            tableHtml.append("</tr>");
                        }
                    }
                    tableHtml.append("</table><br><br><br><p>Thanks<br>Bharath M.</p>");

                    String tableHtmlString = tableHtml.toString().replace("\n", "\\n");

                    // Inject HTML into email body
                    WebElement bodyElement = driver.findElement(By.xpath("//*[@aria-label='Message body']"));
                    ((JavascriptExecutor) driver).executeScript("arguments[0].innerHTML = arguments[1];", bodyElement, tableHtmlString);

                    // Rename and attach the Excel file
                    String newFileName = "SLA_Miss_Report_" + new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date()) + ".xls";
                    File newFile = new File(downloadDir, newFileName);
            		latestFile.renameTo(newFile);
            		
            		String newdownloadedFilePath = newFile.getAbsolutePath();

            		WebElement fileInput = driver.findElement(By.xpath("//input[@type='file']"));
            		fileInput.sendKeys(newdownloadedFilePath);

                    // Wait and send email
                    Thread.sleep(10000);
                    // Uncomment to send the email
                    // driver.findElement(By.xpath("//*[@aria-label='Send']")).click();
                }
            } else {
                System.out.println("No files found in the download directory.");
            }
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }
}
