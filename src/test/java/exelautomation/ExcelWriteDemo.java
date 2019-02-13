package exelautomation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.FileAssert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;

public class ExcelWriteDemo {


@Test
    public void writeExcel() throws Exception{

    String filePath  ="src/test/resources1/country.xls";

    FileInputStream in = new FileInputStream(filePath);
    Workbook workbook = WorkbookFactory.create(in);
    Sheet workSheet = workbook.getSheetAt(0);

    Cell column = workSheet.getRow(0).createCell(2);
    column.setCellValue("Continent");

    Cell continent1 = workSheet.getRow(1).createCell(2);
    continent1.setCellValue("North America");

    FileOutputStream out = new FileOutputStream(filePath);

        workbook.write(out);

}







}
