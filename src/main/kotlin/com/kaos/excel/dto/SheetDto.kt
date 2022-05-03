package com.kaos.excel.dto

// sheet information
data class SheetDto (
    val name: String, // sheet name
    val isRowTitle: Boolean, // whether title is in row or col
    val data: HashMap<String, Array<String>>, // data written in excel file
    val titleOption: SheetTitleDto?, // sheet table title option
    val arrangeOption: SheetArrangeDto?, // sheet table arrange option
    val borderOption: SheetBorderDto? // data decoration option
)