package ApachePOI;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

public class _04_ApachePOIGetAllData {
    public static void main(String[] args) throws IOException {


        String path="src/test/java/ApachePOI/resource/ApacheExcel2.xlsx";



        FileInputStream inputStream=new FileInputStream(path);
        Workbook workbook= WorkbookFactory.create(inputStream);
        Sheet sheet=workbook.getSheet("Sheet1");

        //calisma sayfasindaki toplam satır sayısını veriyor
        int rowCount=sheet.getPhysicalNumberOfRows();//satir sayısı

        for (int i = 0; i <rowCount ; i++) {

            Row row= sheet.getRow(i);//satir
            int cellCount =row.getPhysicalNumberOfCells();//hücre sayısı


            for (int j = 0; j <cellCount ; j++) { //i.satırdaki hücre sayısı kadar dönecek
                Cell hucre=row.getCell(j); //bu satırdaki sıradaki hücreyi aldım
                System.out.print(hucre+" ");
            }
            System.out.println();
        }



    }
}
