package com.navblue.ntracking.excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {
    public void read(String filePath) {
        File srcFile = new File(filePath);
        File[] files = srcFile.listFiles();
        if(files==null || files.length==0) {
            System.out.println("there is no files in path: " + filePath);
            return;
        }
        FileInputStream fis;
        Workbook workbook;
        File reportFile;
        try {
            for(File file : files) {
                //change file format, csv to xls
                reportFile = new File(file.getParentFile().getPath() + "\\" + file.getName().replace("csv", "xls"));
                file.renameTo(reportFile);
                fis = new FileInputStream(reportFile);
                workbook = new XSSFWorkbook(fis);
                Sheet dataSheet = workbook.getSheetAt(0);
                Iterator<Row> iterator = dataSheet.iterator();
                while (iterator.hasNext()) {
                    Row currentRow = iterator.next();
                    Iterator<Cell> cellIterator = currentRow.iterator();
                    while (cellIterator.hasNext()) {
                        Cell currentCell = cellIterator.next();
                        System.out.println(currentCell.getStringCellValue() + "--");
                    }
                }
            }
        }catch (FileNotFoundException e) {
            e.printStackTrace();
        }catch (IOException e) {
            e.printStackTrace();
        }
        /*
         * FileInputStream excelFile = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();
        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                    System.out.print(currentCell.getStringCellValue() + "--");
                } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                    System.out.print(currentCell.getNumericCellValue() + "--");
                }
            }
            System.out.println();
        }
        */
    }
}
