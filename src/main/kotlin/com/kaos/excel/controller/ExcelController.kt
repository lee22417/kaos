package com.kaos.excel.controller

import com.kaos.excel.dto.ExcelInputDto
import com.kaos.excel.services.ExcelGenerater
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestBody
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController

/*
Request Example

{
    "fileName":"test",
    "data":
        {
            "a":"apple",
            "b":2007
        },
    "isRowTitle":true
}

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
            // decide whether title is in row or col
            val isRowTitle: Boolean = excelDto.isRowTitle
            // generate and save excel file
            excelGenerater.generateReport(filePath, excelDto.data, isRowTitle)
            // return ok
            return ResponseEntity.ok("OK")
    }
}