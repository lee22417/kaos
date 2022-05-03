package com.kaos.excel.services

import com.kaos.excel.dto.SheetBorderDto
import com.kaos.excel.dto.SheetTitleDto
import org.apache.commons.codec.binary.Hex
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.PropertyTemplate
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service

@Service
class ExcelCompoentVisualizar {
    fun setTitleFont(workBook: Workbook, titleOption: SheetTitleDto): XSSFFont  {
        val font: XSSFFont = workBook.createFont() as XSSFFont

        // font color
        if(titleOption.fontColor != null) {
            font.setColor(XSSFColor(Hex.decodeHex(titleOption.fontColor), null))
        }
        // size
        if(titleOption.fontSize != null) {
            font.fontHeightInPoints = titleOption.fontSize.toShort()
        }
        // shape :TODO - not working (Mac)
        if(titleOption.italic != null) {
            font.italic = titleOption.italic
        }else if(titleOption.bold != null) {
            font.bold = titleOption.bold
        }
        return font
    }

    // set title cell style
    fun setTitleCellStyle(workBook: Workbook, titleOption: SheetTitleDto): CellStyle {
        val cellStyle = workBook.createCellStyle() as XSSFCellStyle

        // background color
        if(titleOption.backgroundColor != null) {
            // set color
            cellStyle.setFillForegroundColor(XSSFColor(Hex.decodeHex(titleOption.backgroundColor), null))
            // set background pattern
            cellStyle.fillPattern = FillPatternType.SOLID_FOREGROUND
        }
        return cellStyle
    }

    // set table border
    fun setPropertyTemplates(sheet: Sheet, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int, borderOption: SheetBorderDto) {
        var pt: PropertyTemplate = PropertyTemplate()

        // inside border
        if(borderOption.inside == true) {
            pt.drawBorders(
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                BorderStyle.THIN,
                BorderExtent.INSIDE
            )
        }

        // outside border
        if(borderOption.outside == true) {
            pt.drawBorders(
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                BorderStyle.MEDIUM,
                BorderExtent.OUTSIDE
            )
        }

        // apply border
        pt.applyBorders(sheet)
    }

    // set column width
    fun setAutoSizeColumn(sheet: Sheet, firstCol: Int, lastCol: Int) {
        for (colNum in firstCol until lastCol) {
            sheet.autoSizeColumn(colNum)
        }
    }
}