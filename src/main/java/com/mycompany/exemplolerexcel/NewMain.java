/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.exemplolerexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author marce
 */
public class NewMain {
    private static final String fileName = "livros.xls";
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException {
        // TODO code application logic here
        try {
            FileInputStream arquivo = new FileInputStream(new File(NewMain.fileName));
            HSSFWorkbook workbook = new HSSFWorkbook(arquivo);
            HSSFSheet sheetAlunos = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheetAlunos.iterator();
            int linha = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                if( linha > 0){
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getColumnIndex()) {
                        case 0:
                            System.out.println("Código "+cell.getNumericCellValue());
                            break;
                        case 1:
                            System.out.println("Titulo "+cell.getStringCellValue());
                            break;

                        }
                    }
                }
                linha++;
            }
            arquivo.close();

         } catch (FileNotFoundException e) {
                e.printStackTrace();
                System.out.println("Arquivo Excel não encontrado!");
         }
    }
    
}
