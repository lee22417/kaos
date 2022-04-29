package com.kaos.excel.services

import com.kaos.excel.dto.SheetTitleDto
import org.apache.commons.codec.binary.Hex
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service

@Service
class ExcelCompoentVisualizar {
    fun setTitleFont(workBook: XSSFWorkbook, titleOption: SheetTitleDto): XSSFFont  {
        val font: XSSFFont = workBook.createFont()

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

    fun setTitleCellStyle(workBook: XSSFWorkbook, titleOption: SheetTitleDto): XSSFCellStyle {
        val cellStyle: XSSFCellStyle = workBook.createCellStyle()

        // background color
        if(titleOption.backgroundColor != null) {
            // set background pattern
            cellStyle.fillPattern = FillPatternType.SOLID_FOREGROUND
            // set color
            cellStyle.setFillForegroundColor(XSSFColor(Hex.decodeHex(titleOption.backgroundColor), null))
        }

        return cellStyle
    }
}