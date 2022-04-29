package com.kaos.excel.controller

import com.kaos.excel.dto.ExcelInputDto
import com.kaos.excel.services.ExcelGenerater
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestBody
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController


/*
    Request Example : kaos/example/excel
 */

@RestController
@RequestMapping("/excel")
class ExcelController (
    private val excelGenerater: ExcelGenerater
    ){
        @PostMapping()
        fun getExcelReport(@RequestBody excelDto:ExcelInputDto ): ResponseEntity<String> {
            // set file path as desktop
            val filePath: String = System.getProperty("user.home") + "/Desktop/" + excelDto.fileName + ".xlsx"

            // generate and save excel file
            excelGenerater.generateReport(filePath, excelDto.sheet)
            // return ok
            return ResponseEntity.ok("OK")
    }
}