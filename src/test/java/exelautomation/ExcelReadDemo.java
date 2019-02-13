package exelautomation;

import org.apache.poi.ss.usermodel.*;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class ExcelReadDemo {


    @Test

    public void readXLSfile() throws Exception{

        String path = "/Users/zarinakargaeva/Desktop/Countries.xls";

        FileInputStream fileInputStream = new FileInputStream(path);


    Workbook workbook = WorkbookFactory.create(fileInputStream);

    Sheet worksheet  = workbook.getSheetAt(0);

        Row row = worksheet.getRow(0);

        Cell cell1 = row.getCell(0);

        Cell cell2 = row.getCell(1);

        // print cell values

        System.out.println(cell1);
        System.out.println(cell2.toString());


        String country1 = worksheet.getRow(1).getCell(0).toString();
        String capital1 = workbook.getSheetAt(0).getRow(1).getCell(1).toString();

        System.out.println("Country 1 : "+ country1);
        System.out.println("Capital 1 : "+ capital1);

        int countRows = worksheet.getLastRowNum();
        System.out.println(countRows);
        System.out.println();



        for (int i=1; i<= countRows; i++ ){
            //System.out.print("Country # "+ i + " "  +worksheet.getRow(i).getCell(0).toString());
            //System.out.println("  capital is  " +worksheet.getRow(i).getCell(1).toString());

        }

        Map map = new HashMap <String, String> ();

        for (int i=1; i<= countRows; i++ ) {
            String country = worksheet.getRow(i).getCell(0).toString();
            String capital = worksheet.getRow(i).getCell(1).toString();

            map.put(country,capital);
        }
        System.out.println(map);


        workbook.close();
    fileInputStream.close();





    }




}
