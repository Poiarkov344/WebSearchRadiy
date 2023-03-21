import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;

import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


import javax.xml.crypto.Data;
import java.io.*;

import java.io.FileOutputStream;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    public static void main(String[] args) throws InterruptedException, IOException {


        Scanner scan = new Scanner(System.in);
        System.out.println("Enter Код ДК: ");

        ArrayList<String> listDK = new ArrayList<String>();

//        String codeDK = scan.nextLine();
//        String[] codeArray = codeDK.split("\\s+");
//        for (String code : codeArray) {
//            listDK.add(code);
//        }


        boolean value =true;
        while(value){
            String codeDK = scan.nextLine();
            if(codeDK.equals("q")){
                value = false;
            }else{
                listDK.add(codeDK);
            }

        }
        for(String s : listDK){
            System.out.println(s);
        }

        //Creating Excel

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Sheet1");


        //Creating names for a table in Excel

        int numRows = 0;
        int numCols = 8;
        int count =1;
//
//        for(int i =0 ; i<=numRows;++i){
//            XSSFRow row = sheet.createRow(i);
//            for(int j =0; j < numCols;j++){
//                XSSFCell cell = row.createCell(j);
//            }
//        }
        for (int i = 0; i <= numRows; i++) {
             XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < numCols; j++) {
                XSSFCell cell = row.createCell(j);
                switch (j) {
                    case 0 -> cell.setCellValue("№ п/п");
                    case 1 -> cell.setCellValue("Код ДК");
                    case 2 -> cell.setCellValue("Назва предмету закупівлі");
                    case 3 -> cell.setCellValue("Найменування Замовника");
                    case 4 -> cell.setCellValue("Дата оприлюднення");
                    case 5 -> cell.setCellValue("Кінцевий строк подання тендерних пропозицій");
                    case 6 -> cell.setCellValue("Очікувана вартість");
                    case 7 -> cell.setCellValue("Посилання");

                }
            }
        }



//        for(String codDK : listDK) {


            //Path to WebDriver
            System.setProperty("webdriver.chrome.driver", "/Users/yaroslavpoyarkov/Desktop/chromedriver");


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

            wait = new WebDriverWait(driver, Duration.ofSeconds(2));
            wait_filter = new WebDriverWait(driver, Duration.ofSeconds(5));


            //Scrolling

            JavascriptExecutor js = (JavascriptExecutor) driver;

//      search bar
            WebElement searchBar = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section[1]/div/div/div/div/form/input"));

            // Filter



            //List of ДК codes, so we can look for several number of elements

//      search element
//            searchBar.sendKeys(codDK);
//            actions.keyDown(Keys.ENTER).keyUp(Keys.ENTER).perform();

//            Thread.sleep(200);

            wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("search")));
            WebElement status = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/div[1]/label[1]")));
            status.click();

            Thread.sleep(200);

            WebElement filter = wait_filter.until(ExpectedConditions.elementToBeClickable(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/div[1]/div[1]/div[1]/ul[1]/li[2]")));
            filter.click();

            Thread.sleep(200);

            WebElement moreFilter = driver.findElement(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[8]/label[1]"));

            moreFilter.click();

            Thread.sleep(200);

            WebElement priceButton = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section[1]/div/div/div/div/div/div/div[1]/div[8]"));

            priceButton.click();

            Thread.sleep(250);
            WebElement priceInput = driver.findElement(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[8]/div[1]/div[1]/div[1]/label[1]/input[1]"));


            priceInput.sendKeys("200000");

//            Thread.sleep(200);

            WebElement priceConferm = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section[1]/div/div/div/div/div/div/div[1]/div[8]/div/div/div/div/button"));

            priceConferm.click();

            Thread.sleep(600);




            int pages = 0;
            Thread.sleep(500);
            try {

                //Count of pages
                WebElement NumberOfPages = driver.findElement(By.className("paginate"));
                List<WebElement> Pages = NumberOfPages.findElements(By.className("paginate__visible--desktop"));

                for (WebElement row : Pages) {
                    String text = row.getText();
                    if (text.equals("...") || text.equals("")) {

                    } else {
                        pages = Integer.parseInt(text);
                        System.out.println(pages);
                    }
                }
            }catch (NoSuchElementException e){
                pages=1;
                System.out.println(pages);
            }


            Thread.sleep(100);
            //Search every page search result
            Page:
            for (int i = 0; i < pages; i++) {

                Thread.sleep(2000);

                ////      list of search result
                WebElement list = driver.findElement(By.xpath("/html[1]/body[1]/main[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/ul[1]"));
                List<WebElement> listRows1 = list.findElements(By.cssSelector("a"));

                //Count of Codes



                //function

                Link:
                for (WebElement row : listRows1) {
                    numCols = 0;


                    System.out.println(row.getText());


                    String name = row.getText();

                    //Click on each link
                    row.click();


//                    Thread.sleep(100);

                    //List of all Paragraphs, so we can look for the code in them
                    WebElement d = driver.findElement(By.cssSelector("div.tender--description"));
                    List<WebElement> p = d.findElements(By.tagName("p"));


                    //Pattern for looking dor a code
                    Pattern pattern = Pattern.compile("\\d{8}-\\d");
                    String match = "";


                    for (WebElement rowP : p) {
                        Matcher matcher = pattern.matcher(rowP.getText());
                        if (matcher.find()) {
                            match = matcher.group();
                        }
                    }

                    for(int k =0; k<listDK.size();++k) {
                        if(match.equals(listDK.get(k))){
                            System.out.println("Match Approved !!");
                            System.out.println(match);
//                        Thread.sleep(100);
                            //Printing from the new row
                            XSSFRow ExcelRow = sheet.createRow(++numRows);

                            //Printing "Назва предмету закупівлі" in to the cell
                            XSSFCell cell = ExcelRow.createCell(numCols++);

                            cell.setCellValue(count++);
                            cell = ExcelRow.createCell(numCols++);
                            cell.setCellValue(listDK.get(k));
                            cell = ExcelRow.createCell(numCols++);
                            cell.setCellValue(name);


                            // table 1

                            WebElement table1 = driver.findElement(By.cssSelector(".tender--customer.margin-bottom"));
                            List<WebElement> rows4 = table1.findElements(By.className("col-sm-4"));


                            //Looking for "Найменування" in the first table
                            for (WebElement rows : rows4) {
                                if (rows.getText().equals("Найменування:")) {

                                    cell = ExcelRow.createCell(numCols++);

                                    WebElement Data = driver.findElement(By.className("col-sm-6"));
                                    System.out.println(rows.getText().trim() + " " + Data.getText());

                                    cell.setCellValue(Data.getText().trim());
                                }
                            }

                            // table2

                            try { //doing with try in case there is only one table on tha page

                                //looking for table2
                                WebElement table2 = driver.findElement(By.cssSelector(".col-sm-9.tender--customer--inner.margin-bottom.margin-bottom-more"));

                                //creating two lists for to columns that located in table2

                                List<WebElement> Table2Rows8 = table2.findElements(By.className("col-sm-8"));
                                List<WebElement> Table2rows4 = table2.findElements(By.className("col-sm-4"));

                                //creating two iterators for each column

                                Iterator<WebElement> iter1 = Table2Rows8.iterator();
                                Iterator<WebElement> iter2 = Table2rows4.iterator();


                                //Checking both columns at the same time

                                while (iter1.hasNext() && iter2.hasNext()) {

                                    WebElement element1 = iter1.next();
                                    WebElement element2 = iter2.next();

                                    if (element1.getText().equals("Дата оприлюднення:") || element1.getText().equals("Кінцевий строк подання тендерних пропозицій:") || element1.getText().equals("Очікувана вартість:")) {
                                        cell = ExcelRow.createCell(numCols++);

                                        System.out.println(element1.getText().trim() + " " + element2.getText().trim());

                                        cell.setCellValue(element2.getText().trim());
                                    }
                                }
                            } catch (NoSuchElementException e) {
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
                            continue Link;
                        }else{
                            System.out.println("Match not Approved !!");
                            System.out.println(match + " - " + listDK.get(k));
                        }
                        if(k==listDK.size()-1){
                            driver.navigate().back();
                        }
                    }


                }
                try{

                //going to the next page of  the list

                WebElement next = driver.findElement(By.cssSelector(".paginate__btn.next"));
                Thread.sleep(400);

//                wait1.until(ExpectedConditions.visibilityOf(next));

                next.click();

                }catch (NoSuchElementException e){
                    System.out.println(e);
                }

            }


            //quiting driver
            driver.quit();



//        }

        //getting time for the file

        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH:mm:ss");
        String formattedDateTime = now.format(formatter);


        //saving  and closing the file
//        String ExcelName ="Код ДК:" + "123" + " " + formattedDateTime + ".xlsx";
        String ExcelName = "123.xlsx";
        FileOutputStream out = new FileOutputStream(ExcelName);

        workbook.write(out);

        out.close();

        System.out.println("Excel file created successfully.");



    }
}