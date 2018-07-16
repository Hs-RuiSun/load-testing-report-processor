package com.navblue.ntracking.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class ExcelWriter {
    public void writeToExcel(String filePath, List<List<String>> sheetData) {
        try {
            HSSFWorkbook hwb = new HSSFWorkbook();
            HSSFSheet sheet = hwb.createSheet("new sheet");
            writeHeader(sheet, 0);
            
            for (int k = 1; k <= sheetData.size(); k++) {
                ArrayList<String> rowValue = (ArrayList<String>) sheetData.get(k);
                HSSFRow row = sheet.createRow((short) 0 + k);

                for (int p = 0; p < rowValue.size(); p++) {
                    HSSFCell cell = row.createCell((short) p);
                    if(p == 0) {
                        cell.setCellValue(rowValue.get(p).toString());
                    }
                    else {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(rowValue.get(p));
                    }
                }
            }
            String generatedFileName = "load-testing-report.xls";
            FileOutputStream fileOut = new FileOutputStream(filePath + File.separator + generatedFileName);
            hwb.write(fileOut);
            /*//addFilter
            sheet.setAutoFilter(new CellRangeAddress(1, 4, 0, 5));
            
            //setStyle
            CellStyle style;
            Font headerFont = hwb.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 12);
            style = createBorderedStyle(hwb);
            style.setAlignment(CellStyle.ALIGN_CENTER);
            style.setFont(headerFont);*/
            
            fileOut.close();
            System.out.println(generatedFileName + " has been generated");
        } catch (Exception ex) {
        }
    }
    
    public void writeHeader(HSSFSheet sheet, int rowNumber) {
        HSSFRow header = sheet.createRow(rowNumber);
        //1. setCellValue
        HSSFCell cell = header.createCell(0);
        cell.setCellValue("TestTime");
        cell = header.createCell(1);
        cell.setCellValue("Positions");
        cell = header.createCell(2);
        cell.setCellValue("AverageLatency");
        cell = header.createCell(3);
        cell.setCellValue("PercentOverThreshold");
        cell = header.createCell(4);
        cell.setCellValue("Threshold");
        cell = header.createCell(5);
        cell.setCellValue("PullingRate(Millis)");
        
        //2. setLength
        
    }
}
