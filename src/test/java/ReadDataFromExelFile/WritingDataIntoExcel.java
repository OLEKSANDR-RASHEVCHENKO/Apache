package ReadDataFromExelFile;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WritingDataIntoExcel {
    public static void main(String[] args) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Oleksandr\\IdeaProjects\\Apache\\src\\test\\java\\Data\\data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Data");
        XSSFRow row=sheet.createRow(0);
        row.createCell(0).setCellValue("Java");
        row.createCell(1).setCellValue(19);
        row.createCell(2).setCellValue("Automation");

        XSSFRow row1=sheet.createRow(1);
        row1.createCell(0).setCellValue("Python");
        row1.createCell(1).setCellValue(3);
        row1.createCell(2).setCellValue("Automation");

        XSSFRow row2=sheet.createRow(2);
        row2.createCell(0).setCellValue("C#");
        row2.createCell(1).setCellValue(5);
        row2.createCell(2).setCellValue("Automation");

        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();
        System.out.println("File is created");
    }
}
