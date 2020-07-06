package com.example.excel;

import androidx.appcompat.app.AppCompatActivity;

import android.content.res.AssetManager;
import android.os.Bundle;
import android.widget.Toast;


import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.util.Iterator;


public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        createExcel();
        //commit 1
        //commit 2
        //commit 3
        //commit 3
        //commit 2
        //commit 1
    }

    public void createExcel() {
        //new comentario xd

        try {
            String body[][];
            // AssetManager am= getAssets();
            InputStream is = getResources().openRawResource(R.raw.asset);
            HSSFWorkbook wb = new HSSFWorkbook(is);
            HSSFSheet sh = wb.getSheetAt(0);
            Iterator<Row> filas = sh.iterator();
            Iterator<Cell> celdas;
            Row fila;
            Cell celda;
            while (filas.hasNext()) {
                int i = 0;
                fila = filas.next();
                celdas = fila.cellIterator();

                while (celdas.hasNext()) {
                    celda = celdas.next();
                    i++;
                    switch (celda.getCellType()) {
                        case Cell.CELL_TYPE_BLANK:
                            System.out.println(String.valueOf(i));
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            System.out.println(String.valueOf(i) + celda.getNumericCellValue());
                            break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.println(String.valueOf(i) + celda.getStringCellValue());
                            break;
                    }
                }
            }
        } catch (Exception e) {
            System.out.println(e.toString());
        }

    }
}
