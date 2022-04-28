package com.kaos.excel.dto

data class ExcelInputDto (
    val fileName: String, // excel file path
    val data: HashMap<String, String>, // data written in excel file
    val isRowTitle: Boolean // whether title is row or col
    )