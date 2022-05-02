package com.kaos.excel.services

import com.kaos.excel.dto.SheetArrangeDto
import com.kaos.excel.dto.SheetTitleDto
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service

@Service
class ExcelComponentMaker (
    private val compoentVisualizar: ExcelCompoentVisualizar
    ) {

    // tiable with title in row
    fun RowTitleCompomentMaker(
        workBook: XSSFWorkbook,
        sheet: XSSFSheet,
        data: HashMap<String, Array<String>>,
        titleOption: SheetTitleDto?,
        arrangeOption: SheetArrangeDto?
    ): XSSFSheet{
        val titleRowIdx = arrangeOption?.topGap ?: 0;
        val titleColIdx = arrangeOption?.leftGap ?: 0; // col idx
        var directionIdx = titleColIdx; // col idx
        val titleRow = sheet.createRow(titleRowIdx)
        System.out.println(directionIdx)
        data.forEach {
                k,v ->
            // write title (key)
            val titleCell = titleRow.createCell(directionIdx)
            titleCell.setCellValue(k)
            // apply title option
            if (titleOption != null) {
                applyTitleOption(workBook, titleCell, titleOption)
            }

            // write values
            if (directionIdx == titleColIdx) {
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

    // table with title in column
    fun ColTitleCompomentMaker(
        workBook: XSSFWorkbook,
        sheet: XSSFSheet,
        data: HashMap<String, Array<String>>,
        titleOption: SheetTitleDto?,
        arrangeOption: SheetArrangeDto?
    ): XSSFSheet{
        val titleColIdx = arrangeOption?.leftGap ?: 0;
        var directionIdx: Int = arrangeOption?.topGap ?: 0; // row idx

        data.forEach {
            k,v ->
            // specify row
            val row = sheet.createRow(directionIdx)
            // write title (key)
            val titleCell = row.createCell(titleColIdx)
            titleCell.setCellValue(k)
            // apply title option
            if (titleOption != null) {
                applyTitleOption(workBook, titleCell, titleOption)
            }

            // write values
            v.mapIndexed{
                idx, value -> row.createCell(titleColIdx + idx + 1).setCellValue(value)
            }
            // move row
            directionIdx++

        }
        return sheet
    }

    fun applyTitleOption(workBook: XSSFWorkbook, cell: Cell, titleOption: SheetTitleDto){
        // set cell style depends on titleOption
        val cellStyle: XSSFCellStyle = compoentVisualizar.setTitleCellStyle(workBook, titleOption)

        // set font depends on titleOption
        val font = compoentVisualizar.setTitleFont(workBook, titleOption)
        // set font into style
        cellStyle.setFont(font)

        // set cell style
        cell.cellStyle = cellStyle
    }

}