package com.kaos.excel.dto

// sheet information
data class SheetDto (
    val name: String, // sheet name
    val isRowTitle: Boolean, // whether title is in row or col
    val data: HashMap<String, Array<String>>, // data written in excel file
    val arrangeOption: SheetArrangeDto?, // sheet table arrange option
    val decorationOption: SheetDecorationDto? // data decoration option
)