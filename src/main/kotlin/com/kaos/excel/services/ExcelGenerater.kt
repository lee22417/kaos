package com.kaos.excel.services

import com.kaos.excel.dto.SheetDto
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import java.io.FileOutputStream

@Service
class ExcelGenerater(private val componentMaker: ExcelComponentMaker) {
    // generate file
    public fun generateReport(filePath: String, sheetDto: Array<SheetDto>) {
        // create report
        val report = createReport(sheetDto)
        // save report in file path
        saveReport(filePath, report)
    }

    // create new file
    fun createReport(sheetDto: Array<SheetDto>): Workbook {
        // create XSSF Workbook
        val workBook = XSSFWorkbook()
        // create sheets
        sheetDto.map { sheetInfo -> createSheet(workBook, sheetInfo) }
        // return workbook
        return workBook
    }

    // save created file
    fun saveReport(filePath: String, workBook: Workbook){
        // set file path
        val fileOutputStream = FileOutputStream(filePath)
        // create excel file
        workBook.write(fileOutputStream)
        // close
        fileOutputStream.close()
    }


    // create Sheet in file
    fun createSheet(workBook: Workbook, sheetInfo: SheetDto) {
        // sheet name
        val name = sheetInfo.name
        // decide whether title is in row or col
        val isRowTitle = sheetInfo.isRowTitle
        // data in sheet
        val data = sheetInfo.data

        // options in sheet
        val titleOption = sheetInfo.titleOption
        val arrangeOption = sheetInfo.arrangeOption
        val borderOption = sheetInfo.borderOption

        // create sheet with name
        val sheet = workBook.createSheet(name)
        // create sheet with data written in file
        if(isRowTitle) {
            val sheet = componentMaker.RowTitleCompomentMaker(workBook, sheet, data, titleOption, arrangeOption, borderOption)
        }else {
            val sheet = componentMaker.ColTitleCompomentMaker(workBook, sheet, data, titleOption, arrangeOption, borderOption)
        }
    }
}

