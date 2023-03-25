
import org.apache.poi.ss.formula.atp.Switch;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SearchPage {
    public static void main(String[] args) throws Exception {
        String [] KodDK ={"31700000-3", "31680000-6",
                "31220000-4", "31210000-1" ,"31710000-6" ,
                "38810000-6" , "31214000-9" , "32260000-3" ,
                "31219000-4" , "31211110-2" , "31215000-6" ,
                "32320000-2" , "31200000-8" , "38800000-3" ,
                "38400000-9" , "42900000-5" , "31600000-2" ,
                "32400000-7" ,"32000000-3" , "38420000-5"};
        ArrayList<String> listDK = new ArrayList<>(Arrays.asList(KodDK));

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Sheet1");

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


        CellStyle cellStyleBorderBody = workbook.createCellStyle();

        cellStyleBorderBody.setBorderBottom(BorderStyle.THIN);
        cellStyleBorderBody.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBody.setBorderLeft(BorderStyle.THIN);
        cellStyleBorderBody.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBody.setBorderRight(BorderStyle.THIN);
        cellStyleBorderBody.setRightBorderColor(IndexedColors.BLACK.getIndex());
        cellStyleBorderBody.setBorderTop(BorderStyle.THIN);
        cellStyleBorderBody.setTopBorderColor(IndexedColors.BLACK.getIndex());

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

        LocalDate today = LocalDate.now();
        LocalTime now = LocalTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd MMMM yyyy",  Locale.UK);
        DateTimeFormatter formatter1 = DateTimeFormatter.ofPattern("HH:mm:ss");
        String formattedDate = today.format(formatter);
        String formattedTime = now.format(formatter1);

        int numRows = 0;
        int numCols = 8;





        XSSFRow row = sheet.createRow(numRows);
        for(int r =0; r<=3; r++) {
            row = sheet.createRow(numRows++);
            for (int i = 5; i < numCols; ++i) {
                XSSFCell cell = row.createCell(1);

                if(r==0) {
                    cell = row.createCell(i);
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



        for(int i =0; i<listDK.size();++i){
            for(int j =1; j<=2; j++){
                XSSFCell cell = row.createCell(j);
                if(j==1){
                    cell.setCellValue(listDK.get(i));
                }else{
                    switch (listDK.get(i)){
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
            row = sheet.createRow(numRows++);
        }

        row.setHeightInPoints(40);

            for (int j = 0; j < numCols; j++) {
                XSSFCell cell = row.createCell(j);
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
                        sheet.setColumnWidth(4,18 * 256);
                    }
                    case 5 -> {
                        cell.setCellValue("Кінцевий строк подання тендерних пропозицій");
                        cell.setCellStyle(cellStyleBorderBodyWrap);
                        sheet.setColumnWidth(5,18 * 256);
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



        FileOutputStream out = new FileOutputStream("Торши на _ " +formattedDate + "_" + formattedTime + ".xlsx");

        workbook.write(out);

        out.close();

        System.out.println("Excel file created successfully.");
    }



//    public static void main(String[] args) throws IOException {
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
//    }
}

