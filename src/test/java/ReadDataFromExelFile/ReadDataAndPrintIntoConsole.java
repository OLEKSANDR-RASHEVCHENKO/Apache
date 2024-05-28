package ReadDataFromExelFile;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadDataAndPrintIntoConsole {
    public static void main(String[] args) throws IOException {
        FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Oleksandr\\IdeaProjects\\Apache\\src\\test\\java\\Data\\data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
//        XSSFSheet xssfSheet=workbook.getSheet("Sheet1");//so or so
        XSSFSheet xssfSheet=workbook.getSheetAt(0);
        int totalRows=xssfSheet.getLastRowNum();
        int totalCells=xssfSheet.getRow(1).getLastCellNum();
        System.out.println("Number of rows "+ totalRows);//5
        System.out.println("Number of cells "+ totalCells);//4
        for (int r=0;r<totalRows;r++){
            XSSFRow currentRow=xssfSheet.getRow(r);
            for (int c=0;c<totalCells;c++){
                XSSFCell currentCell=currentRow.getCell(c);
                System.out.print(currentCell.toString()+"\t");
            }
            System.out.println();
        }
        workbook.close();
        fileInputStream.close();

    }
}
