package com.kaos.excel.services

import com.kaos.excel.dto.SheetDto
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import java.io.FileOutputStream

@Service
class ExcelGenerater(private val componentMaker: ExcelComponentMaker) {
    // generate report
    public fun generateReport(filePath: String, sheetDto: Array<SheetDto>) {
        // create report
        val report = createReport(sheetDto)
        // save report in file path
        saveReport(filePath, report)
    }

    // create new excel file
    fun createReport(sheetDto: Array<SheetDto>): XSSFWorkbook {
        // create XSSF Workbook
        val workBook = XSSFWorkbook()
        // create sheets
        sheetDto.map { sheetInfo -> createSheet(workBook, sheetInfo) }
        // return workbook
        return workBook
    }

    // save created excel file
    fun saveReport(filePath: String, workBook: Workbook){
        // set file path
        val fileOutputStream = FileOutputStream(filePath)
        // create excel file
        workBook.write(fileOutputStream)
        // close
        fileOutputStream.close()
    }


    // create Sheet in excel file
    fun createSheet(workBook: XSSFWorkbook, sheetInfo: SheetDto) {
        // sheet name
        val name = sheetInfo.name
        // decide whether title is in row or col
        val isRowTitle = sheetInfo.isRowTitle
        // data in excel sheet
        val data = sheetInfo.data

        // create sheet with name
        val sheet = workBook.createSheet(name)
        // create excel sheet with data written in file - isRowTitle = true
        if(isRowTitle) {
            val sheet = componentMaker.RowTitleCompomentMaker(sheet, data)
        }else {
            val sheet = componentMaker.ColTitleCompomentMaker(sheet, data)
        }
    }
}

