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



import java.io.*;

import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    public static void main(String[] args) throws InterruptedException, IOException {


        String [] KodDK ={"31700000-3", "31680000-6",
                "31220000-4", "31210000-1" ,"31710000-6" ,
                "38810000-6" , "31214000-9" , "32260000-3" ,
                "31219000-4" , "31211110-2" , "31215000-6" ,
                "32320000-2" , "31200000-8" , "38800000-3" ,
                "38400000-9" , "42900000-5" , "31600000-2" ,
                "32400000-7" ,"32000000-3" , "38420000-5"};
        ArrayList<String> listDK = new ArrayList<>(Arrays.asList(KodDK));


        LocalDate today = LocalDate.now();
        LocalTime now = LocalTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd MMMM yyyy",   new Locale("uk"));
        DateTimeFormatter formatter1 = DateTimeFormatter.ofPattern("HH_mm_ss");
        String formattedDate = today.format(formatter);
        String formattedTime = now.format(formatter1);

        //Creating Excel

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Sheet1");



        //Create Style for Header

        CellStyle cellStyleHeader = workbook.createCellStyle();

        cellStyleHeader.setBorderBottom(BorderStyle.THIN);
        cellStyleHeader.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleHeader.setBorderLeft(BorderStyle.THIN);
        cellStyleHeader.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleHeader.setBorderRight(BorderStyle.THIN);
        cellStyleHeader.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleHeader.setBorderTop(BorderStyle.THIN);
        cellStyleHeader.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleHeader.setAlignment(HorizontalAlignment.CENTER);
        cellStyleHeader.setVerticalAlignment(VerticalAlignment.CENTER);

        // Create Style for Body

        CellStyle cellStyleBorderBody = workbook.createCellStyle();

        cellStyleBorderBody.setBorderBottom(BorderStyle.THIN);
        cellStyleBorderBody.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBody.setBorderLeft(BorderStyle.THIN);
        cellStyleBorderBody.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBody.setBorderRight(BorderStyle.THIN);
        cellStyleBorderBody.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBody.setBorderTop(BorderStyle.THIN);
        cellStyleBorderBody.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBody.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleBorderBody.setWrapText(true);


        //Create Style for Header/Body with Wrap
        CellStyle cellStyleBorderBodyWrap = workbook.createCellStyle();

        cellStyleBorderBodyWrap.setBorderBottom(BorderStyle.THIN);
        cellStyleBorderBodyWrap.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBodyWrap.setBorderLeft(BorderStyle.THIN);
        cellStyleBorderBodyWrap.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBodyWrap.setBorderRight(BorderStyle.THIN);
        cellStyleBorderBodyWrap.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBodyWrap.setBorderTop(BorderStyle.THIN);
        cellStyleBorderBodyWrap.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBodyWrap.setAlignment(HorizontalAlignment.CENTER);
        cellStyleBorderBodyWrap.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleBorderBodyWrap.setWrapText(true);


        //Creating names for a table in Excel

        int numRows = 0;
        int numCols = 8;
        int count =1;

        XSSFRow rowExcel = sheet.createRow(numRows);

        for(int r =0; r<=3; r++) {
            rowExcel = sheet.createRow(numRows++);
            for (int i = 5; i < numCols; ++i) {
                XSSFCell cell = rowExcel.createCell(1);

                if(r==0) {
                    cell = rowExcel.createCell(i);
                    switch (i) {
                        case 5 -> cell.setCellValue("Зфомовано:");
                        case 6 -> cell.setCellValue(formattedDate + " року");
                        case 7 -> cell.setCellValue(formattedTime);
                    }
                }else{
                    switch (r){
                        case 1 -> cell.setCellValue("Оголошення про Тендери знайдені на сайті https://prozorro.gov.ua зі статусом: \"Подання пропозицій\". Сумма торгів більше 200 тисяч гривень.");
                        case 2 -> cell.setCellValue("За кодами ДК:");
                    }
                }
            }
        }


        for (String s : listDK) {
            for (int j = 1; j <= 2; j++) {
                XSSFCell cell = rowExcel.createCell(j);
                if (j == 1) {
                    cell.setCellValue(s);
                } else {
                    switch (s) {
                        case "31700000-3" -> cell.setCellValue("Електронне, електромеханічне та електротехнічне обладнання");
                        case "31680000-6" -> cell.setCellValue("Електричне приладдя та супутні товари до електричного обладнання");
                        case "31220000-4" -> cell.setCellValue("Елементи електричних схем");
                        case "31210000-1" -> cell.setCellValue("Електрична апаратура для комутування та захисту електричних кіл");
                        case "31710000-6" -> cell.setCellValue("Електронне обладнання");
                        case "38810000-6" -> cell.setCellValue("Обладнання для керування виробничими процесами");
                        case "31214000-9" -> cell.setCellValue("Розподільні пристрої");
                        case "32260000-3" -> cell.setCellValue("Обладнання для передавання даних");
                        case "31219000-4" -> cell.setCellValue("Захисні коробки");
                        case "31211110-2" -> cell.setCellValue("Щити керування");
                        case "31215000-6" -> cell.setCellValue("Обмежувачі напруги");
                        case "32320000-2" -> cell.setCellValue("Телевізійне й аудіовізуальне обладнання");
                        case "31200000-8" -> cell.setCellValue("Електророзподільна та контрольна апаратура");
                        case "38800000-3" -> cell.setCellValue("Обладнання для керування виробничими процесами та пристрої дистанційного керування");
                        case "38400000-9" -> cell.setCellValue("Прилади для перевірки фізичних характеристик");
                        case "42900000-5" -> cell.setCellValue("Універсальні та спеціалізовані машини різні");
                        case "31600000-2" -> cell.setCellValue("Електричні обладнання та апаратура");
                        case "32400000-7" -> cell.setCellValue("Мережі");
                        case "32000000-3" -> cell.setCellValue("Радіо-, телевізійна, комунікаційна, телекомунікаційна та супутня апаратура й обладнання");
                        case "38420000-5" -> cell.setCellValue("Прилади для вимірювання витрати, рівня та тиску рідин і газів");


                    }
                }

            }
            rowExcel = sheet.createRow(numRows++);
        }
        rowExcel = sheet.createRow(numRows++);
        rowExcel.setHeightInPoints(40);





            for (int j = 0; j < numCols; j++) {
                XSSFCell cell = rowExcel.createCell(j);
                cell.setCellStyle(cellStyleHeader);
                switch (j) {
                    case 0 -> {
                        cell.setCellValue("№");
                        sheet.setColumnWidth(0, 3 * 256);
                    }
                    case 1 -> {
                        cell.setCellValue("Код ДК");
                        sheet.setColumnWidth(1, 11 * 256);
                    }
                    case 2 -> {
                        cell.setCellValue("Назва предмету закупівлі");
                        sheet.setColumnWidth(2,52 * 256);
                    }
                    case 3 -> {
                        cell.setCellValue("Найменування Замовника");
                        sheet.setColumnWidth(3,40 * 256);

                    }
                    case 4 -> {
                        cell.setCellValue("Дата оприлюднення");
                        sheet.setColumnWidth(4,19 * 256);
                    }
                    case 5 -> {
                        cell.setCellValue("Кінцевий строк подання тендерних пропозицій");
                        cell.setCellStyle(cellStyleBorderBodyWrap);
                        sheet.setColumnWidth(5,19 * 256);
                    }
                    case 6 -> {
                        cell.setCellValue("Очікувана вартість");
                        sheet.setColumnWidth(6,22 * 256);
                    }
                    case 7 -> {
                        cell.setCellValue("Посилання");
                        sheet.setColumnWidth(7,25 * 256);
                    }
                }
            }







            //Path to Desktop on Windows
//            String desktopPath = System.getProperty("user.home") + "\\Desktop";


            //Path to WebDriver
            System.setProperty("webdriver.chrome.driver", "/Users/yaroslavpoyarkov/Desktop/chromedriver");

//        System.setProperty("webdriver.chrome.driver", desktopPath + "\\chromedriver.exe");


            //Setting up option for WebDriver to be open remotely
            ChromeOptions options = new ChromeOptions();
            options.addArguments("--remote-allow-origins=*");


            //Creating WebDriver and applying option
            WebDriver driver;
            driver = new ChromeDriver(options);


            //Opening link and setting up size of the window (full screen )
            driver.get("https://prozorro.gov.ua");
            driver.manage().window().maximize();




        WebElement status = driver.findElement(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/div[1]/label[1]"));
            status.click();

            Thread.sleep(200);

            WebElement filter = driver.findElement(By.xpath("/html[1]/body[1]/main[1]/div[2]/section[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[5]/div[1]/div[1]/div[1]/ul[1]/li[2]"));
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

            WebElement priceConfirm = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section[1]/div/div/div/div/div/div/div[1]/div[8]/div/div/div/div/button"));

            priceConfirm.click();

            Thread.sleep(600);




            int pages = 0;
            Thread.sleep(500);
            try {

                //Count of pages
                WebElement NumberOfPages = driver.findElement(By.className("paginate"));
                List<WebElement> Pages = NumberOfPages.findElements(By.className("paginate__visible--desktop"));

                for (WebElement row : Pages) {
                    String text = row.getText();
                    try{
                        pages = Integer.parseInt(text);
                        pages=2;
                    }catch(NumberFormatException e){
                        System.out.println(text);
                    }
                }
            }catch (NoSuchElementException e){
                pages=1;
            }


            Thread.sleep(100);
            //Search every page search result PAGE
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

                    //Назва предмету закупівлі
                    String name = row.getText();

                    //Click on each link
                    row.click();



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

                            //Printing from the new row
                            rowExcel = sheet.createRow(numRows++);

                            //start from new row
                            XSSFCell cell = rowExcel.createCell(numCols++);

                            //Print number and move to the next column
                            cell.setCellValue(count++);
                            cell.setCellStyle(cellStyleBorderBodyWrap);
                            cell = rowExcel.createCell(numCols++);


                            //Print KodDK and mover to the next column
                            cell.setCellValue(listDK.get(k));
                            cell.setCellStyle(cellStyleBorderBody);
                            cell = rowExcel.createCell(numCols++);


                            //Set Назва предмету закупівлі
                            cell.setCellValue(name);
                            cell.setCellStyle(cellStyleBorderBody);


                            // table 1

                            WebElement table1 = driver.findElement(By.cssSelector(".tender--customer.margin-bottom"));
                            List<WebElement> rows4 = table1.findElements(By.className("col-sm-4"));


                            //Looking for "Найменування" in the first table
                            for (WebElement rows : rows4) {
                                if (rows.getText().equals("Найменування:")) {

                                    cell = rowExcel.createCell(numCols++);

                                    WebElement Data = driver.findElement(By.className("col-sm-6"));
                                    System.out.println(rows.getText().trim() + " " + Data.getText());

                                    cell.setCellValue(Data.getText().trim());
                                    cell.setCellStyle(cellStyleBorderBody);
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
                                        cell = rowExcel.createCell(numCols++);

                                        System.out.println(element1.getText().trim() + " " + element2.getText().trim());

                                        cell.setCellValue(element2.getText().trim());
                                        cell.setCellStyle(cellStyleBorderBodyWrap);
                                    }
                                }
                            } catch (NoSuchElementException e) {
                                System.out.println(e + "");
                            }


                            //Getting URL of each page

                            String url = driver.getCurrentUrl();
                            System.out.println(url);

                            cell = rowExcel.createCell(numCols);
                            cell.setCellValue(url);
                            cell.setCellStyle(cellStyleBorderBody);


                            //making hyperlink for URL that been passed in the cell

                            CreationHelper createHelper = workbook.getCreationHelper();
                            Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
                            hyperlink.setAddress(url);
                            cell.setHyperlink(hyperlink);


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


                next.click();

                }catch (NoSuchElementException e){
                    System.out.println(e +"");
                }

            }


            //quiting driver
            driver.quit();



        //saving  and closing the file
        String ExcelName ="Торги на _" + formattedDate + "_" + formattedTime + ".xlsx";
        FileOutputStream out = new FileOutputStream(ExcelName);

        workbook.write(out);

        out.close();

        System.out.println("Excel file created successfully.");



    }
}