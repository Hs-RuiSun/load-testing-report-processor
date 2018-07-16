package com.navblue.ntracking;

import com.navblue.ntracking.excel.CSVReader;
import com.navblue.ntracking.excel.ExcelWriter;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Processor{
    public static void main( String[] args ) throws IOException{
        /*String SRC_FILE_NAME = args[0];
        String DES_FILE_NAME = args[1];*/
        
        String SRC_FILE_NAME = "C:\\Users\\ruby.sun\\Downloads\\load-testing-report";
        String DES_FILE_NAME = "C:\\Users\\ruby.sun\\Downloads\\load-testing-report";
        
        /*ExcelReader excelReader = new ExcelReader();
        excelReader.read(SRC_FILE_NAME);*/
        
        CSVReader csvReader = new CSVReader();
        csvReader.readCSVToExcel(SRC_FILE_NAME);
        /*List<List<String>> sheetData = csvReader.readFromCSV(SRC_FILE_NAME);
        ExcelWriter excelWriter = new ExcelWriter();
        excelWriter.writeToExcel(DES_FILE_NAME, sheetData);*/
    }
}
