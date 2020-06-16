package ru.gothmog.app.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteExcel {
    /*
     * 1. Create single data in excel sheet
     * 2. multiple data row in excel sheet
     */
    public void writeSingleCellData(String filePath) throws IOException {
        /**
         * 1. Create a workbook
         * 2. Create a sheet in above workbook
         * 3. Create a row in above sheet
         * 4. Create a cell in above row
         * 5. Set data inside cell
         */
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("firstSheet");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("First Cell");

        //write workbook on output stream
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        //close stream
        fos.close();
        workbook.close();
    }

    public void writeMultipleCellData(String filePath) throws IOException{
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("firstSheet");
        int [][] dataArray  = getRandomDataArray(5,6);
        for (int i = 0; i < dataArray.length; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < dataArray[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(dataArray[i][j]);
            }
        }
        //write workbook on output stream
        File file = new File(filePath);
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        //close stream
        fos.close();
        workbook.close();
    }

    private int[][] getRandomDataArray(int row, int col) {
        int [][] dataArray  = new int[row][col];
        for (int i = 0; i < dataArray.length; i++) {
            for (int j = 0; j < dataArray[i].length ; j++) {
                dataArray[i][j] = (int) (Math.random()*1000);
            }
        }
        return dataArray;
    }
}
