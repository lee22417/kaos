package com.kaos.excel.services

import com.kaos.excel.dto.SheetArrangeDto
import com.kaos.excel.dto.SheetBorderDto
import com.kaos.excel.dto.SheetTitleDto
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.springframework.stereotype.Service

@Service
class ExcelComponentMaker (
    private val componetVisualizar: ExcelCompoentVisualizar
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

            // apply title option
            if (titleOption != null) {
                applyTitleOption(workBook, sheet, titleCell, k, titleOption)
            } else {
                titleCell.setCellValue(k)
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
        componetVisualizar.setAutoSizeColumn(sheet, titleColIdx, directionIdx - 1)

        // apply border option
        if(borderOption != null) {
            applyBorderOption(
                sheet, titleRowIdx + 1, titleRowIdx + rowIdxSize, titleColIdx, directionIdx - 1, borderOption
            )
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

            // create title cell
            val titleCell = row.createCell(titleColIdx)
            if (titleOption != null) {
                // write title (key) and apply option
                applyTitleOption(workBook, sheet, titleCell, k, titleOption)
            } else {
                // write title (key)
                titleCell.setCellValue(k)
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
        componetVisualizar.setAutoSizeColumn(sheet, titleColIdx, directionIdx - 1)

        // apply border option
        if(borderOption != null) {
            applyBorderOption(
                sheet, titleRowIdx, directionIdx - 1, titleColIdx + 1, titleColIdx + colIdxSize, borderOption
            )

        }

        return sheet
    }

    fun applyTitleOption(workBook: Workbook, sheet: Sheet, cell: Cell, text: String, titleOption: SheetTitleDto){
        // set font depends on titleOption
        val font = componetVisualizar.setTitleFont(workBook, titleOption)
        // write text
        val rt = XSSFRichTextString(text)
        // set font
        rt.applyFont(font)
        // set text in cell
        cell.setCellValue(rt)

        // set cell style depends on titleOption
        val cellStyle = componetVisualizar.setTitleCellStyle(workBook, titleOption)
        // set font into style
        cellStyle.setFont(font)
        // set cell style
        cell.cellStyle = cellStyle
    }

    // apply border option
    fun applyBorderOption (
        sheet: Sheet, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int, borderOption: SheetBorderDto
    ){
        // apply border
        componetVisualizar.setBorderPropertyTemplates(sheet, firstRow, lastRow, firstCol, lastCol, borderOption)
    }
}