import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class TranslatorAutomation {

    public static void main(String[] args) {

        String geckoDriverLocation = "geckodriver";
        System.setProperty("webdriver.gecko.driver", geckoDriverLocation);
        WebDriver driver = new FirefoxDriver();

        // region Opening the avis website
        driver.get("https://www.avis.ca/en/home");
        try {
            // Initializing explicit wait
            WebDriverWait wait = new WebDriverWait(driver, 20);

            String pickUpLocation = "//*[@id=\"PicLoc_value\"]";
            String pickUpLocationVal = "Windsor Airport, Windsor, Ontario, Canada-(YQG)";
            typeIntoInputField(wait, driver, pickUpLocation, pickUpLocationVal);

            String returnDate = "//*[@id=\"to\"]";
            String returnDateVal = "12/31/2023";
            typeIntoInputField(wait, driver, returnDate, returnDateVal);

            String selectMyCarButton = "//*[@id=\"res-home-select-car\"]";
            WebElement elementselectMyCarButton = waitForElementVisible(wait, driver, selectMyCarButton);
            elementselectMyCarButton.click();

            String selectThisLocationButton = "/html/body/div[3]/div[6]/div[1]/footer/div[3]/div/div[1]/div/div/div[2]/div[2]/ul/li[1]/div[2]/a";
            WebElement elementselectThisLocationButton = waitForElementVisible(wait, driver, selectThisLocationButton);
            elementselectThisLocationButton.click();

            String dontWantDiscountButton = "//*[@id=\"bx-element-1526913-RKSShaY\"]/button";
            skippableClick(wait, driver, dontWantDiscountButton, 20);

            // get all Car data by class ID
            List<WebElement> divElements = driver.findElements(By.xpath("//div[@class='row avilablecar available-car-box']"));

            // Process and write car data to Excel
            processCarDataToExcel(divElements);

        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            driver.quit();
        }
        // endregion

        // region Opening the carrentals website
        driver.get("https://www.carrentals.com/");
        try {
            // Initializing explicit wait
            WebDriverWait wait = new WebDriverWait(driver, 20);

            String pickUpLocation = "//*[@id=\"wizard-car-pwa-1\"]/div[1]/div[1]/div/div/div/div/div[2]/div[1]/button";
            WebElement pickUpLocationButton = waitForElementVisible(wait, driver, pickUpLocation);
            pickUpLocationButton.click();

            String pickUpLocationInput = "//*[@id=\"location-field-locn\"]";
            String pickUpLocationVal = "Windsor";
            typeIntoInputField(wait, driver, pickUpLocationInput, pickUpLocationVal);

            String selectLocationOption = "//*[@id=\"location-field-locn-menu\"]/section/div/div[2]/div/ul/li[1]/div/button";
            WebElement selectLocationOptionButton = waitForElementVisible(wait, driver, selectLocationOption);
            selectLocationOptionButton.click();

            String search = "//*[@id=\"wizard-car-pwa-1\"]/div[3]/div[2]/button";
            WebElement searchButton = waitForElementVisible(wait, driver, search);
            searchButton.click();

            // get all Car data by class ID
            WebElement olElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[2]/div[1]/div/main/div[2]/div[2]/div/div[1]/div[2]/div[2]/div[2]/ol")));

            // Locate all the <li> elements within the <ol> element
            List<WebElement> liElements = olElement.findElements(By.xpath(".//li"));

            // Print the text of each <li> element
            for (WebElement liElement : liElements) {
                System.out.println(liElement.getText());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            driver.quit();
        }
        // endregion
    }

    private static void typeIntoInputField(WebDriverWait wait, WebDriver driver, String xpath, String text) {
        // Type the specified text into the input field
//        WebElement element = waitForElementVisible(wait, driver, xpath);
        WebElement element = waitForElementVisible(wait, driver, xpath);

        element.click();
        element.clear();
        element.sendKeys(text);
    }

    private static WebElement waitForElementVisible(WebDriverWait wait, WebDriver driver, String xpath) {
        // Wait until the specified element becomes visible
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
//        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
        WebElement element = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(xpath)));
        return element;
    }

    private static String readElementText(WebDriverWait wait, WebDriver driver, String xpath) {
        // Read and return the text of the specified element
        WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(xpath)));
        return element.getText();
    }

    private static void writeStringToExcel(String multilineString, String fileName, int numOfQuotes) {
        // Failsafe for numOfQuotes
        if (numOfQuotes > 50 || numOfQuotes < 1) {
            numOfQuotes = 5;
        }

        // Write a multiline string to an Excel file
        try (Workbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            // Adding headers
            Row headerRow = sheet.createRow(0);
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("Main Sentence");

            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("Google Translator");

            Cell headerCell3 = headerRow.createCell(2);
            headerCell3.setCellValue("Reverso");

            // Splitting the multiline string into lines
            String[] lines = multilineString.split("\n");

            // Creating a new row for each line
            for (int i = 0; i < numOfQuotes; i++) {
                Row row = sheet.createRow(i + 1); // Starting from the second row after headers
                Cell cell = row.createCell(0);
                cell.setCellValue(lines[i]);
            }

            // Saving the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String> readExcelFile(String fileName) {
        List<String> lines = new ArrayList<>();

        try (Workbook workbook = new HSSFWorkbook(new FileInputStream(fileName))) {
            Sheet sheet = workbook.getSheetAt(0); // Assuming the data is in the first sheet

            // Iterate over rows
            Iterator<Row> rowIterator = sheet.iterator();
            // Skip the first row with headers
            if (rowIterator.hasNext()) {
                rowIterator.next();
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0); // Assuming the data is in the first column

                if (cell != null) {
                    lines.add(cell.getStringCellValue());
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return lines;
    }

    private static void writeListToExcel(List<String> translatedTextList, String fileName, int columnNum) {
        try (FileInputStream fileIn = new FileInputStream(fileName);
             Workbook workbook = new HSSFWorkbook(fileIn)) {

            // Assuming the data is in the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Start from the second row after headers
            int rowIndex = 1;

            // Iterate through the translated text list
            for (String translatedText : translatedTextList) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    row = sheet.createRow(rowIndex);
                }

                // Create or update the cell in the "Google Translator" column (index 1)
                Cell cell = row.createCell(columnNum);
                cell.setCellValue(translatedText);

                rowIndex++;
            }

            // Save the updated workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void skippableClick(WebDriverWait wait, WebDriver driver, String xpath, int timeoutSeconds) {
        try {
            WebElement element = wait
                    .withTimeout(timeoutSeconds, TimeUnit.SECONDS)
                    .ignoring(TimeoutException.class)
                    .until(ExpectedConditions.elementToBeClickable(By.xpath(xpath)));

            if (element != null) {
                element.click();
            } else {
                System.out.println("Element not found within the specified timeout. Skipping click.");
            }
        } catch (TimeoutException e) {
            System.out.println("Element not found within the specified timeout. Skipping click.");
        }
    }

    public static boolean isDisplayed(WebElement element) {
        try {
            if(element.isDisplayed())
                return element.isDisplayed();
        }catch (NoSuchElementException ex) {
            return false;
        }
        return false;
    }

    private static void processCarDataToExcel(List<WebElement> divElements) {
        // Creating a workbook and a sheet
        try (Workbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("CarRentalData");

            // Adding headers
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Vehicle Type");
            headerRow.createCell(1).setCellValue("Vehicle Model");
            headerRow.createCell(2).setCellValue("Number of Passengers");
            headerRow.createCell(3).setCellValue("Number of Large Bags");
            headerRow.createCell(4).setCellValue("Number of Small Bags");
            headerRow.createCell(5).setCellValue("Transmission");
            headerRow.createCell(6).setCellValue("Cost");

            // Adding data to the sheet
            int rowIndex = 1; // Starting from the second row after headers

            // Loop through the div elements
            for (WebElement divElement : divElements) {
                // Extract text content from the div element
                String[] carDataArray = divElement.getText().split("\n");

                // Create a new row for each set of car data
                Row row = sheet.createRow(rowIndex);

                // Add data to each column
                row.createCell(0).setCellValue(carDataArray[0]); // Vehicle Type
                row.createCell(1).setCellValue(carDataArray[1]); // Vehicle Model
                String[] passengerData = carDataArray[2].split(" ");
                row.createCell(2).setCellValue(passengerData[0]); // Number of Passengers
                row.createCell(3).setCellValue(passengerData[1]); // Number of Large Bags
                row.createCell(4).setCellValue(passengerData[2]); // Number of Small Bags
                row.createCell(5).setCellValue(carDataArray[3]); // Transmission
                row.createCell(6).setCellValue(carDataArray[5]); // Cost

                rowIndex++;
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("CarRentalData.xls")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
