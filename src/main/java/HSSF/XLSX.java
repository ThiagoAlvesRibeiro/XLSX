package HSSF;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class XLSX {


    public static void main(String[] args) throws FileNotFoundException {


        /*
        Criação da planilha em si
         */

        Workbook wb = new HSSFWorkbook();
        Sheet st = wb.createSheet();

        Row linha1 = st.createRow(0);
        Row linha2 = st.createRow(1);
        Row linha3 = st.createRow(2);
        Row linha4 = st.createRow(3);
        Row linha5 = st.createRow(4);
        Row linha6 = st.createRow(5);

        //coluna A

        Cell cab_titulo = linha1.createCell(0);
        Cell cab_cliente = linha2.createCell(0);
        Cell cab_local = linha3.createCell(0);
        Cell cab_furo = linha4.createCell(0);
        Cell cab_data = linha5.createCell(0);

        //coluna B

        Cell cliente = linha2.createCell(1);
        Cell local = linha3.createCell(1);
        Cell furo = linha4.createCell(1);
        Cell data_inicial = linha5.createCell(1);
        Cell data_final = linha6.createCell(1);

        //coluna D

        Cell nível_de_agua = linha4.createCell(4);
        Cell na_INICIAL = linha4.createCell(5);
        Cell na_10min = linha5.createCell(5);
        Cell na_24h = linha6.createCell(5);

        //coluna E

        Cell inicial_na = linha4.createCell(5);
        Cell min_10 = linha5.createCell(5);
        Cell hrs_24 = linha6.createCell(5);

        cab_titulo.setCellValue("PERFIL DE SONDAGEM E PERCUSSÃO");

        cab_cliente.setCellValue("CLIENTE: ");
        cab_local.setCellValue("LOCAL: ");
        cab_furo.setCellValue("FURO N° ");
        cab_data.setCellValue("DATA: ");

        data_inicial.setCellValue("INICIO");
        data_final.setCellValue("TERMINO");

        nível_de_agua.setCellValue("NA" +
                "\n(m)");
        na_INICIAL.setCellValue("INICIAL");
        na_10min.setCellValue("10min");
        na_24h.setCellValue("24h");

        Font f1 = wb.createFont();
        f1.setFontName("Arial");
        f1.setFontHeight((short) 250);
        f1.setBold(true);

        Font f2 = wb.createFont();
        f2.setFontName("Arial");
        f2.setBold(true);

        CellStyle estilodotitulo = wb.createCellStyle();
        estilodotitulo.setAlignment(HorizontalAlignment.CENTER);
        estilodotitulo.setFont(f1);

        CellStyle estilodocabecalho = wb.createCellStyle();
        estilodocabecalho.setVerticalAlignment(VerticalAlignment.CENTER);
        estilodocabecalho.setAlignment(HorizontalAlignment.CENTER);
        estilodocabecalho.setFont(f2);

        cab_titulo.setCellStyle(estilodotitulo);
        data_final.setCellStyle(estilodocabecalho);
        data_inicial.setCellStyle(estilodocabecalho);
        nível_de_agua.setCellStyle(estilodocabecalho);

        for (int i = 1; i <= st.getLastRowNum(); i++) {

            Row row = st.getRow(i);
            Cell cell = row.getCell(0);

            if (cell != null)
                cell.setCellStyle(estilodocabecalho);
        }

        for (int i=3; i <= 5; i++){

            Row row = st.getRow(i);
            Cell cell = row.getCell(5);

            if (cell != null)

                cell.setCellStyle(estilodocabecalho);

        }

        st.setColumnWidth(0,10*256);

        st.addMergedRegion(new CellRangeAddress(0,0,0,7));
        st.addMergedRegion(new CellRangeAddress(4,5,0,0));
        st.addMergedRegion(new CellRangeAddress(1,1,1,7));
        st.addMergedRegion(new CellRangeAddress(2,2,1,7));
        st.addMergedRegion(new CellRangeAddress(3,3,1,3));
        st.addMergedRegion(new CellRangeAddress(3,5,4,4));
        st.addMergedRegion(new CellRangeAddress(4,4,2,3));
        st.addMergedRegion(new CellRangeAddress(5,5,2,3));


        OutputStream out = new FileOutputStream("Arquivo.xls");

        try {
            if (out != null){

            wb.write(out);
            wb.close();

            }

        } catch (IOException e) {

            throw new RuntimeException(e);

        }

    }

}
