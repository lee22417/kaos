package com.kaos.excel.services

import com.kaos.excel.dto.SheetArrangeDto
import com.kaos.excel.dto.SheetBorderDto
import com.kaos.excel.dto.SheetTitleDto
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.PropertyTemplate
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service

@Service
class ExcelComponentMaker (
    private val compoentVisualizar: ExcelCompoentVisualizar
    ) {

    // tiable with title in row
    fun RowTitleCompomentMaker(
        workBook: Workbook,
        sheet: Sheet,
        data: HashMap<String, Array<String>>,
        titleOption: SheetTitleDto?,
        arrangeOption: SheetArrangeDto?,
        borderOption: SheetBorderDto?
    ): Sheet {
        val titleRowIdx = arrangeOption?.topGap ?: 0;
        val titleColIdx = arrangeOption?.leftGap ?: 0; // col idx
        var rowIdxSize = 0;
        var directionIdx = titleColIdx; // col idx
        val titleRow = sheet.createRow(titleRowIdx)

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
            v.mapIndexed{
                    idx, value ->
                if(directionIdx == titleColIdx ) {
                    val cell = sheet.createRow(titleRowIdx + idx + 1)
                    cell.createCell(directionIdx).setCellValue(value)
                }else {
                    val cell = sheet.getRow(titleRowIdx + idx + 1)
                    cell.createCell(directionIdx).setCellValue(value)
                }
            }

            // save end col
            if(rowIdxSize == 0) {
                rowIdxSize = v.size
            }
            // move col
            directionIdx++
        }

        // column auto size
        compoentVisualizar.setAutoSizeColumn(sheet, titleColIdx, directionIdx - 1)

        // apply border option
        if(borderOption != null) {
            applyBorderOption(sheet, titleRowIdx + 1, titleRowIdx + rowIdxSize, titleColIdx, directionIdx - 1, borderOption)
        }
        return sheet
    }

    // table with title in column
    fun ColTitleCompomentMaker(
        workBook: Workbook,
        sheet: Sheet,
        data: HashMap<String, Array<String>>,
        titleOption: SheetTitleDto?,
        arrangeOption: SheetArrangeDto?,
        borderOption: SheetBorderDto?
    ): Sheet{
        val titleColIdx = arrangeOption?.leftGap ?: 0;
        val titleRowIdx = arrangeOption?.topGap ?: 0;
        var colIdxSize = 0;
        var directionIdx: Int = titleRowIdx // row idx

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

            // save end col
            if(colIdxSize == 0) {
                colIdxSize = v.size
            }
            // move row
            directionIdx++
        }

        // column auto size
        compoentVisualizar.setAutoSizeColumn(sheet, titleColIdx, directionIdx - 1)

        // apply border option
        if(borderOption != null) {
            applyBorderOption(sheet, titleRowIdx, directionIdx - 1, titleColIdx + 1, titleColIdx + colIdxSize, borderOption)
        }

        return sheet
    }

    fun applyTitleOption(workBook: Workbook, cell: Cell, titleOption: SheetTitleDto){
        // set cell style depends on titleOption
        val cellStyle = compoentVisualizar.setTitleCellStyle(workBook, titleOption)

        // set font depends on titleOption
        val font = compoentVisualizar.setTitleFont(workBook, titleOption)
        // set font into style
        cellStyle.setFont(font)

        // set cell style
        cell.cellStyle = cellStyle
    }

    fun applyBorderOption (sheet: Sheet, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int, borderOption: SheetBorderDto){
        compoentVisualizar.setPropertyTemplates(sheet, firstRow, lastRow, firstCol, lastCol, borderOption)
    }


}