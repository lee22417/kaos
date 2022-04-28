package com.kaos.excel.services

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.springframework.stereotype.Service

@Service
class ExcelComponentMaker {
    fun RowTitleCompomentMaker(sheet: XSSFSheet, data: HashMap<String, String>): XSSFSheet{
        val titleColIdx = 0;
        var directionIdx: Int = 0; // row idx
        data.forEach {
            k,v ->
            System.out.println(k + " " + v)
            // specify row
            val row = sheet.createRow(directionIdx)
            // write title (key)
            row.createCell(titleColIdx).setCellValue(k)
            // write value
            row.createCell(1).setCellValue(v)
            //
            directionIdx++

        }
        return sheet
    }

    fun ColTitleCompomentMaker(sheet: XSSFSheet, data: HashMap<String, String>): XSSFSheet{
        val titleRowIdx = 0;
        var directionIdx: Int = 0; // col idx
        val titleRow = sheet.createRow(titleRowIdx)
        data.forEach {
                k,v ->
            System.out.println(k + " " + v)
            // write title (key)
            titleRow.createCell(directionIdx).setCellValue(k)
            // write value
            if (directionIdx == 0) {
                sheet.createRow(1).createCell(directionIdx).setCellValue(v)
            } else {
                sheet.getRow(1).createCell(directionIdx).setCellValue(v)
            }
            //
            directionIdx++

        }
        return sheet
    }
}