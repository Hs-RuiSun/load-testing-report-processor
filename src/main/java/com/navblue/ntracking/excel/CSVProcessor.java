package com.navblue.ntracking.excel;

import com.navblue.ntracking.util.ExcelConfig;
import com.navblue.ntracking.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

public class CSVProcessor {
    /**
     * convert multiple csv files into one excel with multiple sheets
     * that is to say, one csv file will be converted into one sheet
     * @param filePath
     * @throws IOException
     */
    public void convertCSVsToExcel(File srcFile) throws IOException {
        File[] allFiles = srcFile.listFiles();
        String currentLine;
        int sheetIndex = 1;
        
        HSSFWorkbook workbook = new HSSFWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        
        //the final report sheet
        HSSFSheet finalReportSheet = workbook.createSheet("final report");
        //add header to final report
        Row header = finalReportSheet.createRow(0);
        Cell cell = header.createCell(0);
        cell.setCellValue("Positions");
        cell = header.createCell(1);
        cell.setCellValue("PullingRate(Millis)");
        cell = header.createCell(2);
        cell.setCellValue("Threshold");
        cell = header.createCell(3);
        cell.setCellValue("AverageLatency");
        cell = header.createCell(4);
        cell.setCellValue("PercentOverThreshold");      
        
        try {
            //for each csv file
            for(File file : allFiles) {
                HSSFSheet sheet = workbook.createSheet();
                
                int rowIndex = 0;
                BufferedReader bufferedReader = new BufferedReader(new FileReader(file));
                
                while ((currentLine = bufferedReader.readLine()) != null) {
                    HSSFRow row = sheet.createRow(rowIndex++);
                    String cellValueArray[] = currentLine.split(",");
                    for (int cellIndex = 0; cellIndex < cellValueArray.length; cellIndex++) {
                        cell = row.createCell(cellIndex);
                        switch(cellIndex){
                            case 0:
                            case 1:
                            case 2:
                                cell.setCellValue(Integer.parseInt(cellValueArray[cellIndex]));
                                break;
                            case 3:
                            case 4:
                                cell.setCellValue(Double.parseDouble(cellValueArray[cellIndex]));
                                break;
                        }
                    }
                }
                
                int formulaIndex = sheet.getLastRowNum()+1;
                
                /**
                 * in the final report sheet add a new row
                 */
                Row newRow = finalReportSheet.createRow(finalReportSheet.getLastRowNum()+1);
                for(int i=0;i<ExcelConfig.columnNum;i++) {
                    cell = newRow.createCell(i);
                    char columnName = ExcelUtil.getNameByColumnNo(i);
                    //AVERAGE('20180717-001608'!A1:A10)
                    cell.setCellFormula("AVERAGE('Sheet" + sheetIndex + "'!" + columnName + 1 +":" + columnName + (formulaIndex-1) + ")");
                }
                sheetIndex++;
            }
        }catch(Exception e) {
            e.printStackTrace();
        }
        
        String generatedFileName = "load-testing-report.xls";
        FileOutputStream fileOut = new FileOutputStream(srcFile.getParentFile().getAbsolutePath() + File.separator + generatedFileName);
        workbook.write(fileOut);
    }
    
}
