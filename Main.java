package com.soufian;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;

public class Main {

    public static void main(String[] args) throws IOException {

        try (InputStream fs = new FileInputStream("FILE_PATH")) {
            // Read the file 
            Workbook wb = new XSSFWorkbook(fs);
            Sheet sheet = wb.getSheetAt(0);
            
            // Setting Background color
            CellStyle style = wb.createCellStyle();
            style.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
            style.setFillPattern(FillPatternType.BIG_SPOTS);

            for (int i = sheet.getFirstRowNum()+1; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                for (int j = row.getFirstCellNum()+1; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    Row next_row = sheet.getRow(i+1);
                    Cell next_cell = next_row.getCell(j);
                    Cell SGID = row.getCell(2);
                    Cell SGID_next = next_row.getCell(2);
                    if(SGID.getStringCellValue().equals(SGID_next.getStringCellValue())) {

                        System.out.println("past : "+SGID+" actual : "+SGID_next);
                        switch (cell.getCellType()) {
                            case STRING -> {
                                if (!cell.getStringCellValue().equals(next_cell.getStringCellValue())) {
                                    cell.setCellStyle(style);
                                }
                            }
                            case NUMERIC -> {

                                if (!(cell.getNumericCellValue() == next_cell.getNumericCellValue())) {
                                    cell.setCellStyle(style);
                                }
                            }
                            case BOOLEAN -> {

                                if (!(cell.getBooleanCellValue() == next_cell.getBooleanCellValue())) {
                                    cell.setCellStyle(style);
                                }
                            }
                        }

                    }

                }
            }
            OutputStream fileOut = new FileOutputStream("ColoredFile.xls");
            wb.write(fileOut);
            fileOut.close();
        }catch(Exception e) {
            System.out.println(e);
        }

    }

}



