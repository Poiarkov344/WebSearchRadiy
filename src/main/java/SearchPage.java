
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileOutputStream;

public class SearchPage {
    public static void main(String[] args) throws Exception{
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Sheet1");

        XSSFRow row = sheet.createRow(0);

        XSSFCell cell = row.createCell(0);

        cell.setCellValue("Hello, World!");

        FileOutputStream out = new FileOutputStream("output.xlsx");

        workbook.write(out);

        out.close();

        System.out.println("Excel file created successfully.");

    }

    private String column1;
    private String column2;

    public  SearchPage (String column1, String column2){
        this.column1 =column1;
        this.column2 =column2;
    }

}
