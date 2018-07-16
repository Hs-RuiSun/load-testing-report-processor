package com.navblue.ntracking.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class CSVReader {
    
    public void readCSVToExcel(String filePath) throws IOException {
        File[] allFiles = new File(filePath).listFiles();
        String currentLine;
        int sheetIndex = 0;
        HSSFWorkbook hwb = new HSSFWorkbook();
        for(File file : allFiles) {
            int rowIndex = 0;
            String fileName = file.getName().substring(0, file.getName().lastIndexOf("."));
            HSSFSheet sheet = hwb.createSheet(fileName);
            try {
                BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
                while ((currentLine = bufferedReader.readLine()) != null) {
                    HSSFRow row = sheet.createRow(rowIndex++);
                    String cellValueArray[] = currentLine.split(",");
                    for (int p = 0; p < cellValueArray.length; p++) {
                        HSSFCell cell = row.createCell(p);
                        switch(p){
                            case 0:
                                cell.setCellValue(cellValueArray[p]);
                                break;
                            case 1:
                                cell.setCellValue(Integer.parseInt(cellValueArray[p]));
                                break;
                            case 2:
                                cell.setCellValue(Double.parseDouble(cellValueArray[p]));
                                break;
                            case 3:
                                cell.setCellValue(Double.parseDouble(cellValueArray[p]));
                                break;
                            case 4:
                                cell.setCellValue(Integer.parseInt(cellValueArray[p]));
                                break;
                            case 5:
                                cell.setCellValue(Integer.parseInt(cellValueArray[p]));
                                break;
                        }
                    }
                }
                int averageRowIndex = rowIndex - 1;
                HSSFRow row = sheet.createRow(averageRowIndex);
                HSSFCell cell = row.createCell(1);
                cell.setCellFormula("AVERAGE(B" + 1 +":B" + (averageRowIndex-1) + ")");
                cell = row.createCell(2);
                cell.setCellFormula("AVERAGE(C" + 1 +":C" + (averageRowIndex-1) + ")");
                cell = row.createCell(3);
                cell.setCellFormula("AVERAGE(D" + 1 +":D" + (averageRowIndex-1) + ")");
                cell = row.createCell(4);
                cell.setCellFormula("AVERAGE(E" + 1 +":E" + (averageRowIndex-1) + ")");
                cell = row.createCell(5);
                cell.setCellType(CellType.NUMERIC);
                cell.setCellFormula("AVERAGE(F" + 1 +":F" + (averageRowIndex-1) + ")");
                
            }catch(Exception e) {
                e.printStackTrace();
            }
            sheetIndex++;
        }
        String generatedFileName = "load-testing-report.xls";
        FileOutputStream fileOut = new FileOutputStream(filePath + File.separator + generatedFileName);
        hwb.write(fileOut);
    }
    
    public List<List<String>> readFromCSV(String filePath) {
        List<List<String>> sheetList = new ArrayList<List<String>>();
        List<String> rowList = null;
        String currentLine;
        File srcFile = new File(filePath);
        File[] allFiles = srcFile.listFiles();
        for(File file : allFiles) {
            int positionNo = 0;
            long latency = 0;
            double percent = 0;
            int rowCount = 0;
            String currentTime = "";
            int threshold = 0;
            int pullingRate = 0;
            try {
                BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
                while ((currentLine = bufferedReader.readLine()) != null) {
                    rowList = new ArrayList<String>();
                    String cellValueArray[] = currentLine.split(",");
                    currentTime = cellValueArray[0];
                    positionNo = positionNo + Integer.parseInt(cellValueArray[1]);
                    latency = latency + Long.parseLong(cellValueArray[2]);
                    percent = percent + Double.parseDouble(cellValueArray[3]);
                    threshold = Integer.parseInt(cellValueArray[4]);
                    pullingRate = Integer.parseInt(cellValueArray[5]);
                    rowCount++;
                }
                String[] rowValues = {currentTime, positionNo/rowCount+"", latency/rowCount+"", percent/rowCount+"%", threshold+"", pullingRate+""};
                rowList = Arrays.asList(rowValues);
                sheetList.add(rowList);
            }catch(Exception e) {
                e.printStackTrace();
            }
        }
        return sheetList;
    }
}
