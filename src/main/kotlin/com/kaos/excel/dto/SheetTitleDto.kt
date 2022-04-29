package com.kaos.excel.dto

data class SheetTitleDto (
    val fontColor: String?, // title color
    val fontSize: Int?, // title font size
    val bold: Boolean?, // set title font as bold
    val italic: Boolean?, // set title font as italic
    val backgroundColor: String?, // title cell background color
)