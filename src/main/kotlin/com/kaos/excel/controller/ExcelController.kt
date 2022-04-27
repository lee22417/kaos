package com.kaos.excel.controller

import com.kaos.excel.dto.ExcelInputDto
import com.kaos.excel.services.ExcelGenerater
import org.springframework.http.ResponseEntity
import org.springframework.web.bind.annotation.GetMapping
import org.springframework.web.bind.annotation.PostMapping
import org.springframework.web.bind.annotation.RequestBody
import org.springframework.web.bind.annotation.RequestMapping
import org.springframework.web.bind.annotation.RestController

@RestController
@RequestMapping("/excel")
class ExcelController (
    private val excelGenerater: ExcelGenerater
    ){
        @PostMapping()
        fun getExcelReport(@RequestBody excelDto:ExcelInputDto ): ResponseEntity<String> {
            val filePath = System.getProperty("user.home") + "/Desktop/" + excelDto.fileName + ".xlsx"
            System.out.println(filePath)
            excelGenerater.generateReport(filePath)
            return ResponseEntity.ok("OK")
    }
}