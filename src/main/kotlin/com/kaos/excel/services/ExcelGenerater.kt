package com.kaos.excel.services

import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.springframework.stereotype.Service
import java.io.FileOutputStream

@Service
class ExcelGenerater {
    // generate report
    public fun generateReport(fileName: String) {
        // create report
        val report = createReport()
        // save report
        saveReport(fileName, report)
    }

    // create new report
    fun createReport(): XSSFWorkbook {
        // create XSSF Workbook
        val workBook = XSSFWorkbook()
        // create excel sheet
        val sheet = workBook.createSheet()
        return workBook
    }

    // save created Report
    fun saveReport(fileName: String, workBook: Workbook){
        System.out.println(fileName)
        val fileOutputStream = FileOutputStream(fileName)
        workBook.write(fileOutputStream)
        fileOutputStream.close()
    }
}