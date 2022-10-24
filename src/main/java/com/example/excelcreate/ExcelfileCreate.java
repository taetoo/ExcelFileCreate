package com.example.excelcreate;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

@Slf4j
public class ExcelfileCreate {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFCellStyle defaultStyle = workbook.createCellStyle();

        // 테두리 설정
        defaultStyle.setBorderTop(BorderStyle.THIN);
        defaultStyle.setBorderLeft(BorderStyle.THIN);
        defaultStyle.setBorderRight(BorderStyle.THIN);
        defaultStyle.setBorderBottom(BorderStyle.THIN);

        // 줄 바꿈 및 중앙 정렬
        defaultStyle.setWrapText(true);
        defaultStyle.setAlignment(HorizontalAlignment.CENTER);
        defaultStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 시트 생성 및 셀 높이 설정
        XSSFSheet sheet = workbook.createSheet();
        sheet.setDefaultRowHeightInPoints(30);

        // 열의 숫자
        for(int i=0; i < 10; i++){
            Row row = sheet.createRow(i);

            // 헹의 숫자
            for(int j = 0; j < 5; j++){
                Cell cell = row.createCell(j);
                cell.setCellStyle(defaultStyle);
                cell.setCellValue("셀 생성(" + i + "," + j + ")");

                sheet.setColumnWidth(j, 3000);
            }
        }
        try{
            File xlsxFile = new File("/Users/taehyeonkim/Desktop/test.xlsx");
            FileOutputStream fileOut = new FileOutputStream(xlsxFile);
            workbook.write(fileOut);
        } catch (FileNotFoundException e){
            e.printStackTrace();
        }
        finally {
            workbook.close();
        }


    }
}
