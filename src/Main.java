import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.safari.SafariDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;



import javax.lang.model.element.Element;
import java.io.IOException;
import java.lang.ref.SoftReference;
import java.time.Duration;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
//
//        WebDriver driver;
//        driver = new SafariDriver();



        System.setProperty("webdriver.chrome.driver","/Users/yaroslavpoyarkov/Desktop/chromedriver");

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");

        WebDriver driver;
        driver = new ChromeDriver(options);


        Actions actions = new Actions(driver);


        driver.get("https://prozorro.gov.ua");
        driver.manage().window().maximize();
//        driver.get("https://prozorro.gov.ua/tender/UA-2021-12-01-001605-c");
//        driver.get("https://prozorro.gov.ua/tender/UA-2021-12-17-016336-c");
//        driver.get("https://prozorro.gov.ua/search/tender?text=31210000-1");


        //Waiter


        WebDriverWait wait;
        wait = new WebDriverWait(driver, Duration.ofSeconds(2));



        //Scrolling

        JavascriptExecutor js = (JavascriptExecutor) driver;

//         search bar
        WebElement searchBar = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section[1]/div/div/div/div/form/input"));

        // Filter


//         search element
        searchBar.sendKeys("31210000-1");
        actions.keyDown(Keys.ENTER).keyUp(Keys.ENTER).perform();
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("search")));
        WebElement status = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section/div/div/div/div/div/div/div[1]/div[5]/div/label"));
        wait.until(ExpectedConditions.visibilityOf(status));
        status.click();
        WebElement filter = driver.findElement(By.xpath("//*[@id=\"app\"]/div[2]/section[1]/div/div/div/div/div/div/div[1]/div[5]/div/div/div/ul/li[2]"));
        wait.until(ExpectedConditions.visibilityOf(filter));
        filter.click();





//



        //wait for all elements to appear
        wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("/html[1]/body[1]/main[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]")));




//      list
        WebElement list = driver.findElement(By.xpath("/html[1]/body[1]/main[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/ul[1]"));
        List<WebElement> listRows1 = list.findElements(By.cssSelector("a"));




//
//
        // function
        for(WebElement row : listRows1){
            System.out.println(row.getText());
            row.click();
            wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector("body")));
            // table 1
            WebElement table1 = driver.findElement(By.cssSelector(".tender--customer.margin-bottom"));
            List <WebElement> rows4 = table1.findElements(By.className("col-sm-4"));
            for(WebElement rows : rows4){
                if(rows.getText().equals("Найменування:")){
                    WebElement Data = driver.findElement(By.className("col-sm-6"));
                    System.out.println(rows.getText().trim() + " " + Data.getText());
                }
            }

            // table2
            try{
                WebElement table2 = driver.findElement(By.cssSelector(".col-sm-9.tender--customer--inner.margin-bottom.margin-bottom-more"));
            List<WebElement> rows8 = table2.findElements(By.className("col-sm-8"));
            for(WebElement rows : rows8){
                WebElement Data = driver.findElement(By.className("date"));
                if(rows.getText().equals("Дата оприлюднення:")|| rows.getText().equals("Кінцевий строк подання тендерних пропозицій:")){
                    System.out.println(rows.getText().trim() + " " + Data.getText().trim());
                    }
                }
            }catch (NoSuchElementException e){
                System.out.println(e);
            }


            String url = driver.getCurrentUrl();
            System.out.println(url);


            driver.navigate().back();
        }

    }
}