package com.navblue.ntracking;

import com.navblue.ntracking.excel.CSVProcessor;
import java.io.File;
import java.io.IOException;

public class Processor{
    public static void main( String[] args ) throws IOException{
        /*String SRC_FILE_NAME = args[0];
        String DES_FILE_NAME = args[1];*/
        
        String SRC_FILE_NAME = "C:\\Users\\ruby.sun\\Downloads\\load-testing-report";
        CSVProcessor csvReader = new CSVProcessor();
        csvReader.convertCSVsToExcel(new File(SRC_FILE_NAME));
    }
}
