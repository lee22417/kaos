package com.kaos.excel.services

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import java.io.FileOutputStream

@Service
class ExcelGenerater(private val componentMaker: ExcelComponentMaker) {
    // generate report
    public fun generateReport(filePath: String, data: HashMap<String,String>, isRowTitle: Boolean) {
        // create report
        val report = createReport(data, isRowTitle)
        // save report in file path
        saveReport(filePath, report)
    }

    // create new excel file
    fun createReport(data: HashMap<String,String>, isRowTitle: Boolean): XSSFWorkbook {
        // create XSSF Workbook
        val workBook = XSSFWorkbook()
        // create excel sheet with data written in file - isRowTitle = true
        if(isRowTitle == true) {
            val sheet = componentMaker.RowTitleCompomentMaker(workBook.createSheet(), data)
        }else {
            val sheet = componentMaker.ColTitleCompomentMaker(workBook.createSheet(), data)
        }
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
}