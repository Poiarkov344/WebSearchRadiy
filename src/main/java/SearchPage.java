
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileOutputStream;

public class SearchPage {
    public static void main(String[] args) throws Exception{
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Sheet1");

        int numRows = 1;
        int numCols = 7;
        for (int i = 0; i < numRows; i++) {
            XSSFRow row = sheet.createRow(i);
            for (int j = 0; j < numCols; j++) {
                XSSFCell cell = row.createCell(j);
                switch (j) {
                    case 0 -> cell.setCellValue("№ п/п");
                    case 1 -> cell.setCellValue("Назва предмету закупівлі");
                    case 2 -> cell.setCellValue("Найменування");
                    case 3 -> cell.setCellValue("Дата оприлюднення");
                    case 4 -> cell.setCellValue("Кінцевий строк подання тендерних пропозицій");
                    case 5 -> cell.setCellValue("Очікувана вартість");
                    case 6 -> cell.setCellValue("Посилання");
                }
            }
        }


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
