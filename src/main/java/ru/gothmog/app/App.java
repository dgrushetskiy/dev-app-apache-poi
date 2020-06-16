package ru.gothmog.app;


import ru.gothmog.app.write.WriteExcel;

import java.io.IOException;

public class App {
    public static void main( String[] args ) throws IOException {
        WriteExcel writeExcel = new WriteExcel();
        String pathSingle = "/tmp/singleDataFile.xlsx";
        String pathMultiply = "/tmp/multiplyDataFile.xlsx";
        writeExcel.writeSingleCellData(pathSingle);
        writeExcel.writeMultipleCellData(pathMultiply);
        System.out.println("file created");
    }
}
