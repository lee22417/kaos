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
import org.springframework.stereotype.Service
import java.awt.Color

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
        if(titleOption.italic == true) {
            font.italic = true
        }
        // bold :TODO - not working (Mac)
        if(titleOption.bold == true) {
            font.bold = true
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
    fun setBorderPropertyTemplates(
        sheet: Sheet, firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int, borderOption: SheetBorderDto
    ) {
        var pt: PropertyTemplate = PropertyTemplate()
        // set border color
        var insideColor = getColorIndex(borderOption.insideColor)
        var outSideColor = getColorIndex(borderOption.outsideColor)

        // inside border
        if(borderOption.inside == true) {
            pt.drawBorders(
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                BorderStyle.THIN,
                insideColor,
                BorderExtent.INSIDE
            )
        }

        // outside border
        if(borderOption.outside == true) {
            pt.drawBorders(
                CellRangeAddress(firstRow, lastRow, firstCol, lastCol),
                BorderStyle.THIN,
                outSideColor,
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

    // get color index
    fun getColorIndex(color: String?): Short {
        return when (color) {
            "red" -> IndexedColors.RED.index
            "blue" -> IndexedColors.BLUE.index
            "brown" -> IndexedColors.BROWN.index
            "green" -> IndexedColors.GREEN.index
            "yellow" -> IndexedColors.YELLOW.index
            "violet" -> IndexedColors.VIOLET.index
            "coral" -> IndexedColors.CORAL.index
            "lavender" -> IndexedColors.LAVENDER.index
            "sky_blue" -> IndexedColors.SKY_BLUE.index
            "indigo" -> IndexedColors.INDIGO.index
            else -> IndexedColors.BLACK.index
        }
    }
}