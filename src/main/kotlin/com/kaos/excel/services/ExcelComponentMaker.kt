package com.kaos.excel.services

import org.apache.poi.xssf.usermodel.XSSFSheet
import org.springframework.stereotype.Service

@Service
class ExcelComponentMaker {
    fun RowTitleCompomentMaker(sheet: XSSFSheet, data: HashMap<String, Array<String>>): XSSFSheet{
        val titleColIdx = 0;
        var directionIdx: Int = 0; // row idx
        data.forEach {
            k,v ->
            // specify row
            val row = sheet.createRow(directionIdx)
            // write title (key)
            row.createCell(titleColIdx).setCellValue(k)
            // write values
            v.mapIndexed{
                idx, value -> row.createCell(titleColIdx + idx + 1).setCellValue(value)
            }
            // move row
            directionIdx++

        }
        return sheet
    }

    fun ColTitleCompomentMaker(sheet: XSSFSheet, data: HashMap<String, Array<String>>): XSSFSheet{
        val titleRowIdx = 0;
        var directionIdx: Int = 0; // col idx
        val titleRow = sheet.createRow(titleRowIdx)
        data.forEach {
                k,v ->
            // write title (key)
            titleRow.createCell(directionIdx).setCellValue(k)
            // write values
            if (directionIdx == 0) {
                v.mapIndexed{
                    idx, value -> sheet.createRow(titleRowIdx + idx + 1).createCell(directionIdx).setCellValue(value)
                }
            } else {
                v.mapIndexed{
                        idx, value -> sheet.getRow(titleRowIdx + idx + 1).createCell(directionIdx).setCellValue(value)
                }
            }
            // move col
            directionIdx++

        }
        return sheet
    }
}