package ru.gothmog.app;


import ru.gothmog.app.write.WriteExcel;

import java.io.IOException;

public class App {
    public static void main( String[] args ) throws IOException {
        WriteExcel writeExcel = new WriteExcel();
        String path = "/tmp/singleDataFile.xlsx";
        writeExcel.writeSingleCellData(path);
        writeExcel.writeMultipleCellData(path);
        System.out.println("file created");
    }
}
