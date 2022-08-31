package ApachePOI;

import org.apache.poi.ss.usermodel.*;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class _03_ApachePOIStart {
    public static void main(String[] args) throws IOException {

        //dosyanın yolu alındı
        String path="src/test/java/ApachePOI/resource/ApacheExcel2.xlsx";


        //JavaDosya okuma işlemcisine dosyanın yolunu veriyoruz:
        // böylece program ile dosya arasında bağlantı kuruldu
        FileInputStream dosyaOkumaBaglantisi=new FileInputStream(path);


        //Dosya okuma işlemcisi üzerinden çalışma kitabını alıyorum
        //hafızada workbook u alıp olusturdu
        Workbook calismaKitabi= WorkbookFactory.create(dosyaOkumaBaglantisi);

        //istediğim sayfayı alıyorum
        Sheet calismaSayfasi=calismaKitabi.getSheet("Sheet1");

        //istenen satırı alıyorm
        Row satir=calismaSayfasi.getRow(0);

        Cell hucre=satir.getCell(0);

        System.out.println("hucre = " + hucre);
    }
}
