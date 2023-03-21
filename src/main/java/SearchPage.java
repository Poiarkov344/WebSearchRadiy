
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SearchPage {
//    public static void main(String[] args) throws Exception{
//        XSSFWorkbook workbook = new XSSFWorkbook();
//
//        XSSFSheet sheet = workbook.createSheet("Sheet1");
//
//        int numRows = 1;
//        int numCols = 7;
//        for (int i = 0; i < numRows; i++) {
//            XSSFRow row = sheet.createRow(i);
//            for (int j = 0; j < numCols; j++) {
//                XSSFCell cell = row.createCell(j);
//                switch (j) {
//                    case 0 -> cell.setCellValue("№ п/п");
//                    case 1 -> cell.setCellValue("Назва предмету закупівлі");
//                    case 2 -> cell.setCellValue("Найменування");
//                    case 3 -> cell.setCellValue("Дата оприлюднення");
//                    case 4 -> cell.setCellValue("Кінцевий строк подання тендерних пропозицій");
//                    case 5 -> cell.setCellValue("Очікувана вартість");
//                    case 6 -> cell.setCellValue("Посилання");
//                }
//            }
//        }
//
//        LocalDateTime now = LocalDateTime.now();
//        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH:mm:ss");
//        String formattedDateTime = now.format(formatter);
//
//        FileOutputStream out = new FileOutputStream(formattedDateTime +".xlsx");
//
//        workbook.write(out);
//
//        out.close();
//
//        System.out.println("Excel file created successfully.");
//
//    }


    public static void main(String[] args) throws IOException {
//        System.setProperty("webdriver.chrome.driver","/Users/yaroslavpoyarkov/Desktop/chromedriver");
//
//
//        //Setting up option for WebDriver to be open remotely
//        ChromeOptions options = new ChromeOptions();
//        options.addArguments("--remote-allow-origins=*");
//
//
//        //Creating WebDriver and applying option
//        WebDriver driver;
//        driver = new ChromeDriver(options);
//
//        driver.get("https://prozorro.gov.ua/tender/UA-2023-03-08-006806-a");
//
//        WebElement d = driver.findElement(By.cssSelector("div.tender--description"));
////        System.out.println(d.getText());
//        List<WebElement> b =d.findElements(By.tagName("p"));
//
//        Pattern pattern = Pattern.compile("\\d{8}-\\d");
//
//        for(WebElement row : b){
//            Matcher matcher = pattern.matcher(row.getText());
//            String match ="";
//            if(matcher.find()){
//                match = matcher.group();
//                System.out.println(match);
//            }
//            if(match.equals("31210000-1 , 32000000-3")){
//                break;
//            }
////            System.out.println(row.getText());
//        }
        System.setProperty("java.io.tmpdir", "/Users/yaroslavpoyarkov/Desktop/programing/Java/Projects/WebSearchRadiy");

        String inputFile = "123.xlsx";
        String outputFile = "4.xlsx";
        int columnToCheck = 7;




        try(FileInputStream inputStream = new FileInputStream(new File(inputFile));
            Workbook workbook = WorkbookFactory.create(inputStream);
            FileOutputStream outputStream = new FileOutputStream(outputFile)) {
            Sheet sheet = workbook.getSheetAt(0);
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                // Check if the value in the specified column matches the value in the previous row
                if (i > 0 && row.getCell(columnToCheck).toString().equals(sheet.getRow(i - 1).getCell(columnToCheck).toString())) {
                    // If the values match, delete the current row
                    sheet.removeRow(row);
                }

            }
            sheet.shiftRows(1, sheet.getLastRowNum(), -1);
            workbook.write(outputStream);
        }catch(IOException e){
            e.printStackTrace();
        }
    }

    private String column1;
    private String column2;

    public  SearchPage (String column1, String column2){
        this.column1 =column1;
        this.column2 =column2;
    }

}

