import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;

public class ExportData {
    public ExportData() throws IOException {

        //Stworzenie skoryszytu
        Workbook skoroszyt = new XSSFWorkbook();

        //dodanie arkusza
        Sheet arkusz = skoroszyt.createSheet("Nazwa naszego arkusza");

        //Stworzenie w pierwszym wierszu i komurek z nazwami kolumn
        //Stworzenie pierwszego wierza
        Row wiersz = arkusz.createRow(0);
        for (int j = 0; j < 5; j++) {
            //Stworzenie komurek i wpisanie w nie nazwy
            Cell komurka = wiersz.createCell(j);
            komurka.setCellValue("Kolumna "+j);
        }

        Random ran = new Random();
        int r=1;
        while (r<41){
            //Stworzenie kolejnych wierszy
            wiersz = arkusz.createRow(r);
            for (int j = 0; j < 5; j++) {
                //Stworzenie kolejnych komurek z danymi
                Cell komurka = wiersz.createCell(j);
                komurka.setCellValue(ran.nextInt(101));
            }
            r++;
        }

        wiersz = arkusz.createRow(r+1);
        //Stworzenie kolejnych komurek z formułą

        for (int i=0; i<5; i++){
            Cell komurka = wiersz.createCell(0);
            komurka.setCellFormula("SUM("+(char)(i+65)+"2:"+(char)(i+65)+"41)");
        }
        Cell komurka0 = wiersz.createCell(5);
        komurka0.setCellFormula("AVERAGE(A43:E43)");

        //Malowanie według formuły wielkości liczby
        SheetConditionalFormatting formatowanie = arkusz.getSheetConditionalFormatting();

        //Warunek 1 kiedy liczba > 65
        ConditionalFormattingRule zasada1 = formatowanie.createConditionalFormattingRule(ComparisonOperator.GT, "65");
        PatternFormatting wypelnienie1 = zasada1.createPatternFormatting();
        wypelnienie1.setFillBackgroundColor(IndexedColors.RED.index);
        wypelnienie1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        //Warunek w kiedy liczba < 35
        ConditionalFormattingRule zasada2 = formatowanie.createConditionalFormattingRule(ComparisonOperator.LT, "35");
        PatternFormatting wypelnienie2 = zasada2.createPatternFormatting();
        wypelnienie2.setFillBackgroundColor(IndexedColors.GREEN.index);
        wypelnienie2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);


        CellRangeAddress[] obszarFormatowania = {
                CellRangeAddress.valueOf("A2:E41")
        };

        formatowanie.addConditionalFormatting(obszarFormatowania, zasada1, zasada2);

        wiersz = arkusz.createRow(r+1);
        Cell komurka1 = wiersz.createCell(0);
        komurka1.setCellFormula("COUNTIF(A2:E41,\"<35\")");
        komurka1 = wiersz.createCell(1);
        komurka1.setCellFormula("COUNTIF(A2:E41,\">=35\")-COUNTIF(A2:E41,\">65\")");
        komurka1 = wiersz.createCell(2);
        komurka1.setCellFormula("COUNTIF(A2:E41,\">65\")");


        //Ustawienie rozmiaru columny na automat aby dane zawsze się wyświetlały i nie były przycięte
        for(int i = 0; i < 5; i++) {
            arkusz.autoSizeColumn(i);
        }

        //Zapis do pliku razem z mechanizme chroniącym przed nadpisanie
        File fileTest = new File("./NaszArkusz.xlsx");
        boolean yet = false;
        int manyYet = 1;
        while(!yet){
            if (fileTest.exists()){
                fileTest = new File("./NaszArkusz("+manyYet+").xlsx");
            } else {
                yet = true;
            }
            manyYet++;
        }
        FileOutputStream fileOut = new FileOutputStream(fileTest);
        skoroszyt.write(fileOut);
        fileOut.close();
        skoroszyt.close();

        System.out.println("Plik został wygenerowany");
    }
}
