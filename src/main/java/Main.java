import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;

import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


import javax.xml.crypto.Data;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import java.io.FileOutputStream;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;
import java.util.List;

public class Main {
    public static void main(String[] args) throws InterruptedException, IOException {



        //Path to WebDriver
        System.setProperty("webdriver.chrome.driver","/Users/yaroslavpoyarkov/Desktop/chromedriver");


        //Setting up option for WebDriver to be open remotely
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");


        //Creating WebDriver and applying option
        WebDriver driver;
        driver = new ChromeDriver(options);


        //Creating actions for driver
        Actions actions = new Actions(driver);

        //Opening link and setting up size of the window (full screen )
        driver.get("https://prozorro.gov.ua");
        driver.manage().window().maximize();



        //Waiter


        WebDriverWait wait;
        WebDriverWait wait_filter;
        WebDriverWait wait1;
        wait = new WebDriverWait(driver, Duration.ofSeconds(2));
        wait_filter = new WebDriverWait(driver, Duration.ofSeconds(5));
        wait1 = new WebDriverWait(driver, Duration.ofSeconds(60));



        //Scrolling

        JavascriptExecutor js = (JavascriptExecutor) driver;

//      search bar
        WebElement searchBar = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section[1]/div/div/div/div/form/input"));

        // Filter


//      search element
        searchBar.sendKeys("31210000-1");
        actions.keyDown(Keys.ENTER).keyUp(Keys.ENTER).perform();

        Thread.sleep(2500);

        wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("search")));
        WebElement status = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/div[1]/label[1]")));
        status.click();

        Thread.sleep(2000);

        WebElement filter = wait_filter.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/div[1]/div[1]/div[1]/ul[1]/li[2]")));
        filter.click();

        Thread.sleep(2000);


        //Count of pages
        WebElement NumberOfPages = driver.findElement(By.className("paginate"));
        List<WebElement> Pages = NumberOfPages.findElements(By.className("paginate__visible--desktop"));

        int pages =0;
        for(WebElement row : Pages){
            String text = row.getText();
            if(text.equals("...")){

            }else{
                pages = Integer.parseInt(text);
            }

        }


        //Creating Excel

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Sheet1");


        //Creating names for a table in Excel

        int numRows = 0;
        int numCols = 7;
        for (int i = 0; i <= numRows; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < numCols; j++) {
                XSSFCell cell = row.createCell(j);
                switch (j) {
                    case 0 -> cell.setCellValue("№ п/п");
                    case 1 -> cell.setCellValue("Назва предмету закупівлі");
                    case 2 -> cell.setCellValue("Найменування Замовника");
                    case 3 -> cell.setCellValue("Дата оприлюднення");
                    case 4 -> cell.setCellValue("Кінцевий строк подання тендерних пропозицій");
                    case 5 -> cell.setCellValue("Очікувана вартість");
                    case 6 -> cell.setCellValue("Посилання");
                }
            }
        }


        //Search every page search result
        for(int i =0; i<pages; i++){

            Thread.sleep(2000);

            ////      list of search result
            WebElement list = driver.findElement(By.xpath("/html[1]/body[1]/main[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/ul[1]"));
            List<WebElement> listRows1 = list.findElements(By.cssSelector("a"));


            //function

            for(WebElement row : listRows1){
                numCols =1;

                //Printing from the new row
                XSSFRow ExcelRow = sheet.createRow(++numRows);

                System.out.println(row.getText());

                //Printing "Назва предмету закупівлі" in to the cell
                XSSFCell cell = ExcelRow.createCell(numCols++);
                cell.setCellValue(row.getText());

                //Click on each link
                row.click();

                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector("body")));
                // table 1

                WebElement table1 = driver.findElement(By.cssSelector(".tender--customer.margin-bottom"));
                List <WebElement> rows4 = table1.findElements(By.className("col-sm-4"));

                //Looking for "Найменування" in the first table
                for(WebElement rows : rows4){
                    if(rows.getText().equals("Найменування:")){

                        cell = ExcelRow.createCell(numCols++);

                        WebElement Data = driver.findElement(By.className("col-sm-6"));
                        System.out.println(rows.getText().trim() + " " + Data.getText());

                        cell.setCellValue(Data.getText().trim());
                    }
                }

                // table2
                try{ //doing with try in case there is only one table on tha page
                    //looking for table2
                    WebElement table2 = driver.findElement(By.cssSelector(".col-sm-9.tender--customer--inner.margin-bottom.margin-bottom-more"));
                    //creating two lists for to columns that located in table2
                    List<WebElement> Table2Rows8 = table2.findElements(By.className("col-sm-8"));
                    List<WebElement> Table2rows4 = table2.findElements(By.className("col-sm-4"));

                    //creating two iterators for each column
                    Iterator<WebElement> iter1 = Table2Rows8.iterator();
                    Iterator<WebElement> iter2 = Table2rows4.iterator();


                    //Checking both columns at the same time
                    while(iter1.hasNext() && iter2.hasNext()){

                        WebElement element1 = iter1.next();
                        WebElement element2 = iter2.next();

                        if(element1.getText().equals("Дата оприлюднення:")|| element1.getText().equals("Кінцевий строк подання тендерних пропозицій:") || element1.getText().equals("Очікувана вартість:")){
                            cell = ExcelRow.createCell(numCols++);

                            System.out.println(element1.getText().trim() + " " + element2.getText().trim());

                            cell.setCellValue(element2.getText().trim());
                        }
                    }
                }catch (NoSuchElementException e){
                    System.out.println(e);
                }

                //Getting URL of each page
                String url = driver.getCurrentUrl();
                System.out.println(url);

                cell = ExcelRow.createCell(numCols);
                cell.setCellValue(url);

                //making hyperlink for URL that been passed in the cell
                CreationHelper createHelper = workbook.getCreationHelper();
                Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
                hyperlink.setAddress(url);
                cell.setHyperlink((org.apache.poi.ss.usermodel.Hyperlink) hyperlink);


                driver.navigate().back();
            }

            //going to the next page of  the list
            WebElement next = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(".paginate__btn.next")));
            wait1.until(ExpectedConditions.visibilityOf(next));
            next.click();

        }


        //quiting driver
        driver.quit();



        //getting time for the file
        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH:mm:ss");
        String formattedDateTime = now.format(formatter);

        //saving  and closing the file
        FileOutputStream out = new FileOutputStream(formattedDateTime + ".xlsx");

        workbook.write(out);

        out.close();

        System.out.println("Excel file created successfully.");
    }
}